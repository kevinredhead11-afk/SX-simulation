"""
SX Steady-State Simulation
Ref: Alind – Description of SX Steady State Model (D2EHPA / HCl system)

Key equations (from PDF):
  D = P * [HCl]^Q                          (eq 5)
  k_{i,n} = y_{i,n} / x_{i,n}             (distribution coefficient per cell)
  Tridiagonal system solved with TDMA      (eq 8-9)

Circuit cell numbering (1-indexed, left→right):
  1 … n_ext          →  Extraction  (cell 1 = barren organic inlet / raffinate exit)
  n_ext+1 … n_ext+n_scr  →  Scrub
  n_ext+n_scr+1 … N  →  Strip  (cell N = strip acid inlet / loaded organic exit)

Organic flows left→right; aqueous flows right→left.
Feed enters at cell B = n_ext.
Strip acid enters at cell D = N.
"""

import numpy as np


# ---------------------------------------------------------------------------
# Chemistry helpers
# ---------------------------------------------------------------------------

def d_value(P, Q, hcl):
    """D = P * [HCl]^Q.  Clamp HCl to avoid log(0)."""
    return P * max(hcl, 1e-12) ** Q


# ---------------------------------------------------------------------------
# TDMA (Thomas algorithm)
# ---------------------------------------------------------------------------

def tdma(a, b, c, d):
    """
    Solve tridiagonal system.  Vectors are 0-indexed length n.
    a[i]*x[i-1] + b[i]*x[i] + c[i]*x[i+1] = d[i]
    Boundary: a[0] = 0, c[n-1] = 0 (already set by caller).
    """
    n = len(b)
    c_ = np.zeros(n, dtype=float)
    d_ = np.zeros(n, dtype=float)
    x  = np.zeros(n, dtype=float)

    if abs(b[0]) < 1e-300:
        raise ValueError("TDMA: zero pivot at row 0")
    c_[0] = c[0] / b[0]
    d_[0] = d[0] / b[0]

    for i in range(1, n):
        denom = b[i] - a[i] * c_[i-1]
        if abs(denom) < 1e-300:
            raise ValueError(f"TDMA: zero pivot at row {i}")
        c_[i] = c[i] / denom if i < n - 1 else 0.0
        d_[i] = (d[i] - a[i] * d_[i-1]) / denom

    x[-1] = d_[-1]
    for i in range(n - 2, -1, -1):
        x[i] = d_[i] - c_[i] * x[i + 1]
    return x


# ---------------------------------------------------------------------------
# Core SX steady-state solver
# ---------------------------------------------------------------------------

def solve_sx(n_ext, n_scr, n_str,
             hcl_ext, hcl_scr, hcl_str,
             F, S, O, R,
             elements):
    """
    Solve cell-by-cell mass balance for each REE using TDMA.

    Parameters
    ----------
    n_ext, n_scr, n_str : int   – number of stages per section
    hcl_ext/scr/str     : float – HCl normality (N) per section
    F, S, O             : float – Feed, Strip-solution, Organic flow rates
    R                   : float – Reflux fraction (0-1); scrub flow = R*S
    elements            : list of dicts {name, P, Q, feed_conc, mw}

    Returns
    -------
    y_all : dict name→ndarray[N]  organic concentration per cell (g/L)
    x_all : dict name→ndarray[N]  aqueous concentration per cell (g/L)
    """
    N  = n_ext + n_scr + n_str
    B  = n_ext              # feed cell (1-indexed)
    C  = n_ext + n_scr      # last scrub cell
    D  = N                  # last strip cell

    SR = S * R              # scrub flow rate

    y_all = {}
    x_all = {}

    for el in elements:
        P      = el['P']
        Q      = el['Q']
        xi_F   = el['feed_conc']

        # Distribution coefficient per section (constant within section)
        k_e = d_value(P, Q, hcl_ext)   # extraction
        k_s = d_value(P, Q, hcl_scr)   # scrub
        k_t = d_value(P, Q, hcl_str)   # strip

        # Helper: k for cell n (1-indexed)
        def k_of(n):
            if n <= n_ext:
                return k_e
            elif n <= C:
                return k_s
            else:
                return k_t

        a = np.zeros(N, dtype=float)
        b = np.zeros(N, dtype=float)
        c = np.zeros(N, dtype=float)
        d = np.zeros(N, dtype=float)

        for idx in range(N):
            n = idx + 1   # 1-indexed cell

            # ---- Strip section (C+1 … D) ----
            if n > C:
                a[idx] = O
                b[idx] = -(O + S / k_of(n))
                c[idx] = S / k_of(n + 1) if n < D else 0.0
                d[idx] = 0.0

            # ---- Scrub section (B+1 … C) ----
            elif n > B:
                a[idx] = O
                b[idx] = -(O + SR / k_of(n))
                c[idx] = SR / k_of(n + 1)   # n+1 is still in scrub or first strip
                d[idx] = 0.0

            # ---- Extraction section (1 … B) ----
            else:
                if n == 1:
                    # Cell A – barren organic enters; a=0
                    a[idx] = 0.0
                    b[idx] = -(O + (SR + F) / k_of(n))
                    c[idx] = (SR + F) / k_of(n + 1)   # n+1 still extraction (if B>1) or scrub
                    d[idx] = 0.0
                elif n == B:
                    # Cell B – feed enters
                    # Superdiagonal uses scrub k because cell B+1 is in scrub section
                    a[idx] = O
                    b[idx] = -(O + (SR + F) / k_of(n))
                    c[idx] = SR / k_s if n_scr > 0 else 0.0
                    d[idx] = -(xi_F * F)
                else:
                    # Cells A+1 … B-1
                    a[idx] = O
                    b[idx] = -(O + (SR + F) / k_of(n))
                    c[idx] = (SR + F) / k_of(n + 1)
                    d[idx] = 0.0

        # Enforce boundary conditions
        a[0]  = 0.0
        c[-1] = 0.0

        # Special case: if n_ext == 1, Cell A and Cell B are the same cell
        if n_ext == 1:
            b[0] = -(O + (SR + F) / k_e)
            c[0] = SR / k_s if n_scr > 0 else 0.0
            d[0] = -(xi_F * F)

        y = tdma(a, b, c, d)
        y_all[el['name']] = y

        # Aqueous from organic via distribution coefficient
        x = np.array([y[n - 1] / k_of(n) for n in range(1, N + 1)])
        x_all[el['name']] = x

    return y_all, x_all, N, n_ext, n_scr, n_str


# ---------------------------------------------------------------------------
# Post-processing
# ---------------------------------------------------------------------------

def compute_summary(y_all, x_all, N, n_ext, n_scr, n_str, elements, F, S, O, R):
    """
    Raffinate = aqueous leaving cell 1.
    Preg      = aqueous at cell C+1 (index C in 0-based).

    % columns are PURITY (composition) in each stream:
      % Feed     = element_i_feed  / total_REE_feed  * 100
      % Raffinate = element_i_raff  / total_REE_raff  * 100
      % Preg      = element_i_preg  / total_REE_preg  * 100
    """
    C = n_ext + n_scr

    feed_concs = {el['name']: el['feed_conc'] for el in elements}
    raff_concs = {el['name']: max(x_all[el['name']][0], 0.0)  for el in elements}
    preg_concs = {el['name']: max(x_all[el['name']][C], 0.0)  for el in elements}

    total_feed = sum(feed_concs.values()) or 1.0
    total_raff = sum(raff_concs.values()) or 1.0
    total_preg = sum(preg_concs.values()) or 1.0

    results = {}
    for el in elements:
        name = el['name']
        results[name] = {
            'feed':     feed_concs[name],
            'raff':     raff_concs[name],
            'preg':     preg_concs[name],
            'pct_feed': 100.0 * feed_concs[name] / total_feed,
            'pct_raff': 100.0 * raff_concs[name] / total_raff,
            'pct_preg': 100.0 * preg_concs[name] / total_preg,
        }

    loaded_org = sum(max(y_all[el['name']][n_ext - 1], 0.0) for el in elements)
    return results, loaded_org


# ---------------------------------------------------------------------------
# Default element parameters
# ---------------------------------------------------------------------------

DEFAULT_ELEMENTS = [
    {'name': 'Pr', 'P': 0.0008,  'Q': -1.973, 'feed_conc': 0.0,  'mw': 140.9},
    {'name': 'Nd', 'P': 0.0031,  'Q': -2.541, 'feed_conc': 20.8, 'mw': 144.2},
    {'name': 'Tb', 'P': 0.378,   'Q': -2.624, 'feed_conc': 0.0,  'mw': 158.9},
    {'name': 'Dy', 'P': 0.5959,  'Q': -2.431, 'feed_conc': 2.68, 'mw': 162.5},
]


# ---------------------------------------------------------------------------
# GUI  (imports here so tkinter is never loaded when used as a library)
# ---------------------------------------------------------------------------

def _load_gui_deps():
    global tk, ttk, messagebox, plt, COLORS
    import tkinter as _tk
    from tkinter import ttk as _ttk, messagebox as _mb
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as _plt
    tk = _tk
    ttk = _ttk
    messagebox = _mb
    plt = _plt

COLORS = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']


class SXApp:
    def __init__(self):
        super().__init__()
        self.title("SX Simulation by Alind Chandra")
        self.resizable(True, True)
        self._build_ui()

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        # Top row: three input panels side by side
        top = ttk.Frame(main)
        top.pack(fill="x")
        self._build_general_panel(top)
        self._build_pq_panel(top)
        self._build_feed_panel(top)

        # Action row
        mid = ttk.Frame(main)
        mid.pack(pady=8)
        ttk.Button(mid, text="Generate profile", command=self._run).pack(side="left", padx=12)
        self.status_var = tk.StringVar()
        tk.Label(mid, textvariable=self.status_var,
                 fg="#009900", font=("Arial", 10, "bold")).pack(side="left")

        # Output table
        self._build_output_panel(main)

    def _build_general_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="General Inputs", padding=6)
        frame.pack(side="left", fill="y", padx=(0, 10))

        fields = [
            ("Extraction Stages:",    "n_ext",    "10"),
            ("Scrub Stages:",         "n_scr",    "10"),
            ("Strip Stages:",         "n_str",    "5"),
            ("(N) Acid Conc. (Ext):", "hcl_ext",  "0.50"),
            ("(N) Acid Conc. (Scr):", "hcl_scr",  "1.37"),
            ("(N) Acid Conc. (Str):", "hcl_str",  "5"),
            ("Feed Flowrate (F):",    "F",         "25"),
            ("Strip Flowrate (S):",   "S",         "12"),
            ("Organic Flow (O):",     "O",         "32.5"),
        ]
        self.gen_vars = {}
        for row, (label, key, default) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=2)
            var = tk.StringVar(value=default)
            ttk.Entry(frame, textvariable=var, width=10).grid(row=row, column=1, padx=6, pady=2)
            self.gen_vars[key] = var

        # Reflux slider
        r = len(fields)
        ttk.Label(frame, text="Reflux (R) %:").grid(row=r, column=0, sticky="w", pady=(6, 0))
        self.reflux_var = tk.IntVar(value=70)

        sf = ttk.Frame(frame)
        sf.grid(row=r + 1, column=0, columnspan=2, sticky="ew", pady=2)

        self.reflux_disp = tk.Label(sf, text="70", width=4, anchor="e")
        self.reflux_disp.pack(side="right")

        pct_label = tk.Label(sf, text="70%", width=4)
        pct_label.pack(side="right")

        def on_slide(v):
            iv = int(float(v))
            self.reflux_disp.config(text=str(iv))
            pct_label.config(text=f"{iv}%")

        tk.Scale(sf, from_=0, to=100, orient="horizontal",
                 variable=self.reflux_var, showvalue=False,
                 command=on_slide).pack(side="left", fill="x", expand=True)

    def _build_pq_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Distribution Coeffs, P & Q", padding=6)
        frame.pack(side="left", fill="y", padx=(0, 10))

        for col, txt in enumerate(("El", "P", "Q")):
            ttk.Label(frame, text=txt, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=6, pady=2)

        self.pq_vars = []
        for row, el in enumerate(DEFAULT_ELEMENTS):
            ttk.Label(frame, text=el['name']).grid(row=row + 1, column=0, sticky="w", pady=3)
            p_var = tk.StringVar(value=str(el['P']))
            q_var = tk.StringVar(value=str(el['Q']))
            ttk.Entry(frame, textvariable=p_var, width=10).grid(row=row + 1, column=1, padx=4)
            ttk.Entry(frame, textvariable=q_var, width=9).grid(row=row + 1, column=2, padx=4)
            self.pq_vars.append((el['name'], p_var, q_var))

    def _build_feed_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Feed Concentrations, g/L", padding=6)
        frame.pack(side="left", fill="y")

        self.feed_vars = []
        for row, el in enumerate(DEFAULT_ELEMENTS):
            ttk.Label(frame, text=el['name']).grid(row=row, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=str(el['feed_conc']))
            ttk.Entry(frame, textvariable=var, width=10).grid(row=row, column=1, padx=6)
            self.feed_vars.append(var)

    def _build_output_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Compositions & Distributions", padding=6)
        frame.pack(fill="both", expand=True, pady=(4, 0))

        self.loaded_lbl_var = tk.StringVar(value="Loaded Organic, g/L, in stage –: –")
        tk.Label(frame, textvariable=self.loaded_lbl_var,
                 font=("Arial", 9, "bold")).pack(anchor="w", pady=(0, 4))

        row_frame = ttk.Frame(frame)
        row_frame.pack(fill="both", expand=True)

        # Compositions
        cf = ttk.LabelFrame(row_frame, text="Compositions, g/L", padding=4)
        cf.pack(side="left", fill="both", expand=True, padx=(0, 8))
        comp_cols = ("Element", "Feed Conc. (g/L)", "Raff Conc. (g/L)", "Preg Conc. (g/L)")
        self.comp_tree = ttk.Treeview(cf, columns=comp_cols, show="headings", height=5)
        for col in comp_cols:
            self.comp_tree.heading(col, text=col)
            self.comp_tree.column(col, width=130, anchor="center")
        self.comp_tree.pack(fill="both", expand=True)

        # Distributions
        df = ttk.LabelFrame(row_frame, text="Distributions, %", padding=4)
        df.pack(side="left", fill="both", expand=True)
        dist_cols = ("Element", "% Feed", "% Raffinate", "% Preg")
        self.dist_tree = ttk.Treeview(df, columns=dist_cols, show="headings", height=5)
        for col in dist_cols:
            self.dist_tree.heading(col, text=col)
            self.dist_tree.column(col, width=110, anchor="center")
        self.dist_tree.pack(fill="both", expand=True)

    # ------------------------------------------------------------------ Logic

    def _get_inputs(self):
        gv = self.gen_vars
        n_ext   = int(gv['n_ext'].get())
        n_scr   = int(gv['n_scr'].get())
        n_str   = int(gv['n_str'].get())
        hcl_ext = float(gv['hcl_ext'].get())
        hcl_scr = float(gv['hcl_scr'].get())
        hcl_str = float(gv['hcl_str'].get())
        F = float(gv['F'].get())
        S = float(gv['S'].get())
        O = float(gv['O'].get())
        R = self.reflux_var.get() / 100.0

        elements = []
        for i, (name, p_var, q_var) in enumerate(self.pq_vars):
            elements.append({
                'name':      name,
                'P':         float(p_var.get()),
                'Q':         float(q_var.get()),
                'feed_conc': float(self.feed_vars[i].get()),
                'mw':        DEFAULT_ELEMENTS[i]['mw'],
            })
        return n_ext, n_scr, n_str, hcl_ext, hcl_scr, hcl_str, F, S, O, R, elements

    def _run(self):
        self.status_var.set("")
        try:
            args = self._get_inputs()
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
            return

        n_ext, n_scr, n_str, hcl_ext, hcl_scr, hcl_str, F, S, O, R, elements = args

        try:
            y_all, x_all, N, n_ext, n_scr, n_str = solve_sx(
                n_ext, n_scr, n_str,
                hcl_ext, hcl_scr, hcl_str,
                F, S, O, R,
                elements
            )
        except Exception as e:
            messagebox.showerror("Simulation Error", str(e))
            return

        summary, loaded_org = compute_summary(
            y_all, x_all, N, n_ext, n_scr, n_str, elements, F, S, O, R
        )

        # Update label
        self.loaded_lbl_var.set(
            f"Loaded Organic, g/L, in stage {n_ext}: {loaded_org:.2f}"
        )

        # Refresh tables
        for tree in (self.comp_tree, self.dist_tree):
            for row in tree.get_children():
                tree.delete(row)

        for el in elements:
            d = summary[el['name']]
            self.comp_tree.insert("", "end", values=(
                el['name'],
                f"{d['feed']:.2f}",
                f"{d['raff']:.2f}",
                f"{d['preg']:.2f}",
            ))
            self.dist_tree.insert("", "end", values=(
                el['name'],
                f"{d['pct_feed']:.2f}",
                f"{d['pct_raff']:.2f}",
                f"{d['pct_preg']:.2f}",
            ))

        self.status_var.set("Profile generated!")
        self._plot(y_all, x_all, N, n_ext, n_scr)

    def _plot(self, y_all, x_all, N, n_ext, n_scr):
        stages = np.arange(1, N + 1)

        def add_dividers(ax):
            ax.axvline(n_ext + 0.5,        color='gray', ls='--', lw=0.8, alpha=0.7)
            ax.axvline(n_ext + n_scr + 0.5, color='gray', ls='--', lw=0.8, alpha=0.7)

        names = [el['name'] for el in DEFAULT_ELEMENTS
                 if el['name'] in x_all]

        # Figure 1 – Absolute aqueous concentrations
        fig1, ax1 = plt.subplots(figsize=(7, 4.5))
        fig1.canvas.manager.set_window_title("Figure 1")
        ax1.set_title("Aqueous Distributions – Absolute")
        ax1.set_xlabel("Stage")
        ax1.set_ylabel("Aqueous Concentration X, g/L")
        for i, name in enumerate(names):
            ax1.plot(stages, x_all[name], 'o-', color=COLORS[i % len(COLORS)],
                     label=name, markersize=4)
        add_dividers(ax1)
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        fig1.tight_layout()

        # Figure 2 – % of Total aqueous
        fig2, ax2 = plt.subplots(figsize=(7, 4.5))
        fig2.canvas.manager.set_window_title("Figure 2")
        ax2.set_title("Aqueous Distributions – % of Total")
        ax2.set_xlabel("Stage")
        ax2.set_ylabel("Aqueous Distribution, %")

        total_aq = np.zeros(N)
        for name in names:
            total_aq += np.abs(x_all[name])
        total_aq = np.where(total_aq < 1e-12, 1.0, total_aq)

        for i, name in enumerate(names):
            pct = 100.0 * x_all[name] / total_aq
            ax2.plot(stages, pct, 'o-', color=COLORS[i % len(COLORS)],
                     label=name, markersize=4)
        add_dividers(ax2)
        ax2.set_ylim(-5, 105)
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        fig2.tight_layout()

        plt.show()


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    _load_gui_deps()
    app = SXApp()
    app.mainloop()
