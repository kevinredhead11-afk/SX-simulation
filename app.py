"""
SX Steady-State Simulation – Web App (Streamlit)
"""

import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from sx_simulation import solve_sx, compute_summary, DEFAULT_ELEMENTS
import copy

st.set_page_config(page_title="SX Simulation", layout="wide")
st.title("SX Simulation – Steady State Model")
st.markdown("**D2EHPA / HCl system** — Based on Alind Chandra's model")

# ---------------------------------------------------------------------------
# Sidebar – Inputs
# ---------------------------------------------------------------------------
st.sidebar.header("General Inputs")

n_ext = st.sidebar.number_input("Extraction Stages", min_value=1, max_value=50, value=10, step=1)
n_scr = st.sidebar.number_input("Scrub Stages",      min_value=1, max_value=50, value=10, step=1)
n_str = st.sidebar.number_input("Strip Stages",      min_value=1, max_value=50, value=5,  step=1)

st.sidebar.markdown("---")
hcl_ext = st.sidebar.number_input("(N) Acid Conc. Extraction", value=0.50, format="%.3f")
hcl_scr = st.sidebar.number_input("(N) Acid Conc. Scrub",      value=1.37, format="%.3f")
hcl_str = st.sidebar.number_input("(N) Acid Conc. Strip",      value=5.0,  format="%.2f")

st.sidebar.markdown("---")
F = st.sidebar.number_input("Feed Flowrate (F)",   value=25.0,  format="%.1f")
S = st.sidebar.number_input("Strip Flowrate (S)",  value=12.0,  format="%.1f")
O = st.sidebar.number_input("Organic Flow (O)",    value=32.5,  format="%.1f")
R = st.sidebar.slider("Reflux (R) %", min_value=0, max_value=100, value=70) / 100.0

# ---------------------------------------------------------------------------
# Main area – P/Q and Feed concentrations
# ---------------------------------------------------------------------------
col_pq, col_feed = st.columns(2)

with col_pq:
    st.subheader("Distribution Coefficients  P & Q")
    pq_data = []
    cols = st.columns(3)
    cols[0].markdown("**Element**")
    cols[1].markdown("**P**")
    cols[2].markdown("**Q**")
    for el in DEFAULT_ELEMENTS:
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"**{el['name']}**")
        p = c2.number_input(f"P_{el['name']}", value=el['P'], format="%.4f",
                            label_visibility="collapsed", key=f"P_{el['name']}")
        q = c3.number_input(f"Q_{el['name']}", value=el['Q'], format="%.3f",
                            label_visibility="collapsed", key=f"Q_{el['name']}")
        pq_data.append((el['name'], p, q))

with col_feed:
    st.subheader("Feed Concentrations (g/L)")
    feed_data = []
    for el in DEFAULT_ELEMENTS:
        c1, c2 = st.columns([1, 2])
        c1.markdown(f"**{el['name']}**")
        fc = c2.number_input(f"feed_{el['name']}", value=el['feed_conc'], format="%.2f",
                             label_visibility="collapsed", key=f"feed_{el['name']}")
        feed_data.append(fc)

# ---------------------------------------------------------------------------
# Run simulation
# ---------------------------------------------------------------------------
if st.button("▶  Generate Profile", type="primary", use_container_width=True):

    elements = []
    for i, (name, p, q) in enumerate(pq_data):
        elements.append({
            'name':      name,
            'P':         p,
            'Q':         q,
            'feed_conc': feed_data[i],
            'mw':        DEFAULT_ELEMENTS[i]['mw'],
        })

    try:
        y_all, x_all, N, n_ext_r, n_scr_r, n_str_r = solve_sx(
            n_ext, n_scr, n_str,
            hcl_ext, hcl_scr, hcl_str,
            F, S, O, R,
            elements
        )
        summary, loaded_org = compute_summary(
            y_all, x_all, N, n_ext_r, n_scr_r, n_str_r,
            elements, F, S, O, R
        )
    except Exception as e:
        st.error(f"Simulation error: {e}")
        st.stop()

    st.success(f"Profile generated!  —  Loaded Organic at stage {n_ext}: **{loaded_org:.2f} g/L**")

    # ---- Output tables ----
    st.subheader("Compositions & Distributions")
    tcol1, tcol2 = st.columns(2)

    import pandas as pd

    comp_rows = []
    dist_rows = []
    for el in elements:
        d = summary[el['name']]
        comp_rows.append({
            "Element":         el['name'],
            "Feed Conc. (g/L)": f"{d['feed']:.2f}",
            "Raff Conc. (g/L)": f"{max(d['raff'],0):.2f}",
            "Preg Conc. (g/L)": f"{max(d['preg'],0):.2f}",
        })
        dist_rows.append({
            "Element":      el['name'],
            "% Feed":       f"{d['pct_feed']:.2f}",
            "% Raffinate":  f"{max(d['pct_raff'],0):.2f}",
            "% Preg":       f"{max(d['pct_preg'],0):.2f}",
        })

    with tcol1:
        st.markdown("**Compositions, g/L**")
        st.dataframe(pd.DataFrame(comp_rows), hide_index=True, use_container_width=True)

    with tcol2:
        st.markdown("**Distributions, %**")
        st.dataframe(pd.DataFrame(dist_rows), hide_index=True, use_container_width=True)

    # ---- Plots ----
    stages = np.arange(1, N + 1)
    COLORS = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']
    names  = [el['name'] for el in elements]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

    def add_dividers(ax):
        ax.axvline(n_ext + 0.5,        color='gray', ls='--', lw=0.8, alpha=0.7)
        ax.axvline(n_ext + n_scr + 0.5, color='gray', ls='--', lw=0.8, alpha=0.7)

    # Absolute
    ax1.set_title("Aqueous Distributions – Absolute", fontweight='bold')
    ax1.set_xlabel("Stage")
    ax1.set_ylabel("Aqueous Concentration X, g/L")
    for i, name in enumerate(names):
        ax1.plot(stages, x_all[name], 'o-', color=COLORS[i % 4], label=name, markersize=4)
    add_dividers(ax1)
    ax1.legend()
    ax1.grid(True, alpha=0.3)

    # % Total
    total = np.zeros(N)
    for name in names:
        total += np.abs(x_all[name])
    total = np.where(total < 1e-12, 1.0, total)

    ax2.set_title("Aqueous Distributions – % of Total", fontweight='bold')
    ax2.set_xlabel("Stage")
    ax2.set_ylabel("Aqueous Distribution, %")
    for i, name in enumerate(names):
        pct = 100.0 * x_all[name] / total
        ax2.plot(stages, pct, 'o-', color=COLORS[i % 4], label=name, markersize=4)
    add_dividers(ax2)
    ax2.set_ylim(-5, 105)
    ax2.legend()
    ax2.grid(True, alpha=0.3)

    fig.tight_layout()
    st.pyplot(fig)
