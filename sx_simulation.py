"""
SX Steady-State Model – Core math only (no GUI dependencies)
Ref: Alind – Description of SX Steady State Model (D2EHPA / HCl system)

Key equations:
  D = P * [HCl]^Q                    (eq 5)
  k_{i,n} = y_{i,n} / x_{i,n}       (distribution coefficient per cell)
  Tridiagonal system solved with TDMA (eq 8-9)

Circuit cell numbering (1-indexed, left→right):
  1 … n_ext               → Extraction  (cell 1 = barren organic / raffinate exit)
  n_ext+1 … n_ext+n_scr  → Scrub
  n_ext+n_scr+1 … N      → Strip  (cell N = strip acid inlet / preg exit)

Organic flows left→right; aqueous flows right→left.
Feed enters at cell B = n_ext.
Strip acid enters at cell D = N.
"""

import numpy as np


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
    Solve tridiagonal system (0-indexed length n).
    a[i]*x[i-1] + b[i]*x[i] + c[i]*x[i+1] = d[i]
    Boundary: a[0] = 0, c[n-1] = 0 (set by caller).
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
    n_ext, n_scr, n_str : int   – stages per section
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
    B  = n_ext
    C  = n_ext + n_scr
    D  = N
    SR = S * R

    y_all = {}
    x_all = {}

    for el in elements:
        P    = el['P']
        Q    = el['Q']
        xi_F = el['feed_conc']

        k_e = d_value(P, Q, hcl_ext)
        k_s = d_value(P, Q, hcl_scr)
        k_t = d_value(P, Q, hcl_str)

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
            n = idx + 1

            if n > C:                       # Strip section
                a[idx] = O
                b[idx] = -(O + S / k_of(n))
                c[idx] = S / k_of(n + 1) if n < D else 0.0
                d[idx] = 0.0

            elif n > B:                     # Scrub section
                a[idx] = O
                b[idx] = -(O + SR / k_of(n))
                c[idx] = SR / k_of(n + 1)
                d[idx] = 0.0

            else:                           # Extraction section
                if n == 1:
                    a[idx] = 0.0
                    b[idx] = -(O + (SR + F) / k_of(n))
                    c[idx] = (SR + F) / k_of(n + 1)
                    d[idx] = 0.0
                elif n == B:
                    a[idx] = O
                    b[idx] = -(O + (SR + F) / k_of(n))
                    c[idx] = SR / k_s if n_scr > 0 else 0.0
                    d[idx] = -(xi_F * F)
                else:
                    a[idx] = O
                    b[idx] = -(O + (SR + F) / k_of(n))
                    c[idx] = (SR + F) / k_of(n + 1)
                    d[idx] = 0.0

        a[0]  = 0.0
        c[-1] = 0.0

        if n_ext == 1:
            b[0] = -(O + (SR + F) / k_e)
            c[0] = SR / k_s if n_scr > 0 else 0.0
            d[0] = -(xi_F * F)

        y = tdma(a, b, c, d)
        y_all[el['name']] = y
        x_all[el['name']] = np.array([y[n - 1] / k_of(n) for n in range(1, N + 1)])

    return y_all, x_all, N, n_ext, n_scr, n_str


# ---------------------------------------------------------------------------
# Post-processing
# ---------------------------------------------------------------------------

def compute_summary(y_all, x_all, N, n_ext, n_scr, n_str, elements, F, S, O, R):
    """
    % columns = PURITY (composition) in each stream:
      % Feed      = element_i_feed / total_REE_feed * 100
      % Raffinate = element_i_raff / total_REE_raff * 100
      % Preg      = element_i_preg / total_REE_preg * 100
    """
    C = n_ext + n_scr

    feed_concs = {el['name']: el['feed_conc']                    for el in elements}
    raff_concs = {el['name']: max(x_all[el['name']][0],   0.0)   for el in elements}
    preg_concs = {el['name']: max(x_all[el['name']][C],   0.0)   for el in elements}

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
