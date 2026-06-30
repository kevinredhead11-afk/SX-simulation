#!/usr/bin/env python3
"""
Build SX_Steady_State_Model.xlsx
Alind formulation  —  D2EHPA / HCl  —  TDMA steady-state solver
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter as gcl

# ── Layout constants ──────────────────────────────────────────────────────────
N      = 25          # total stages (fixed: 10 ext + 10 scr + 5 str)
R0     = 5           # first data row in SOLVER sheet
R1     = R0 + N - 1  # last data row  = 29

# ── INPUTS sheet cell refs (absolute) ────────────────────────────────────────
I = "INPUTS!"
REF = {
    "n_ext":   f"{I}$B$5",  "n_scr":   f"{I}$B$6",  "n_str":   f"{I}$B$7",
    "He":      f"{I}$B$8",  "Hs":      f"{I}$B$9",  "Ht":      f"{I}$B$10",
    "F":       f"{I}$B$11", "S":       f"{I}$B$12",
    "O":       f"{I}$B$13", "R":       f"{I}$B$14",
    # P, Q, Feed  (rows 18-21 for Pr/Nd/Tb/Dy; col C=P, D=Q; col C for Feed rows 25-28)
    "P_Pr":  f"{I}$C$18", "Q_Pr":  f"{I}$D$18", "XF_Pr": f"{I}$C$25",
    "P_Nd":  f"{I}$C$19", "Q_Nd":  f"{I}$D$19", "XF_Nd": f"{I}$C$26",
    "P_Tb":  f"{I}$C$20", "Q_Tb":  f"{I}$D$20", "XF_Tb": f"{I}$C$27",
    "P_Dy":  f"{I}$C$21", "Q_Dy":  f"{I}$D$21", "XF_Dy": f"{I}$C$28",
}

ELEMS = ["Pr", "Nd", "Tb", "Dy"]

# Col start (1-indexed) for each element's 8-column block in SOLVER
#   offsets: 0=D, 1=a, 2=b, 3=c, 4=d, 5=c', 6=d', 7=y_org
ESTART = {"Pr": 3, "Nd": 11, "Tb": 19, "Dy": 27}

# Aqueous x_i (=y/D) columns  (35=AI, 36=AJ, 37=AK, 38=AL)
XCOL = {"Pr": 35, "Nd": 36, "Tb": 37, "Dy": 38}

# ── Colour palette ────────────────────────────────────────────────────────────
NAVY  = "1F4E79"; BLUE  = "2E75B6"; LBLUE = "BDD7EE"; XBLUE = "EBF3FA"
ORNG  = "ED7D31"; GRN   = "70AD47"; RED   = "C00000"; YLW   = "FFF2CC"
LGRN  = "E2EFDA"; LRED  = "FCE4D6"; WHT   = "FFFFFF"; GRY   = "D9D9D9"

# ── Style helpers ─────────────────────────────────────────────────────────────
def fill(c): return PatternFill("solid", fgColor=c)
def fnt(bold=False, col="000000", sz=11):
    return Font(bold=bold, color=col, size=sz, name="Calibri")
def aln(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
_t = Side(style="thin", color="AAAAAA")
_m = Side(style="medium", color="1F4E79")
def brd(): return Border(left=_t, right=_t, top=_t, bottom=_t)
def brd_med(): return Border(left=_m, right=_m, top=_m, bottom=_m)

def set_cell(ws, row, col, value=None, formula=None, bold=False, col_font=WHT,
             sz=11, bg=None, h="center", wrap=False, num_fmt=None, border=True):
    c = ws.cell(row=row, column=col)
    c.value = formula if formula else value
    c.font = fnt(bold=bold, col=col_font, sz=sz)
    if bg:
        c.fill = fill(bg)
    c.alignment = aln(h=h, wrap=wrap)
    if border:
        c.border = brd()
    if num_fmt:
        c.number_format = num_fmt
    return c

# ── Formula builders ──────────────────────────────────────────────────────────
def f_D(row, elem):
    P, Q = REF[f"P_{elem}"], REF[f"Q_{elem}"]
    ne, ns = REF["n_ext"], REF["n_scr"]
    He, Hs, Ht = REF["He"], REF["Hs"], REF["Ht"]
    return (f"=IF(A{row}<={ne},{P}*{He}^{Q},"
            f"IF(A{row}<={ne}+{ns},{P}*{Hs}^{Q},"
            f"{P}*{Ht}^{Q}))")

def f_a(row):
    return f"=IF(A{row}=1,0,{REF['O']})"

def f_b(row, Dc):
    O, S, R, F = REF["O"], REF["S"], REF["R"], REF["F"]
    ne, ns, D = REF["n_ext"], REF["n_scr"], f"{Dc}{row}"
    return (f"=IF(A{row}>{ne}+{ns},"
            f"-{O}-{S}/{D},"
            f"IF(A{row}>{ne},"
            f"-{O}-{S}*{R}/{D},"
            f"-{O}-({S}*{R}+{F})/{D}))")

def f_c(row, Dc):
    S, R, F = REF["S"], REF["R"], REF["F"]
    ne, ns = REF["n_ext"], REF["n_scr"]
    if row == R1:
        return "=0"
    Dn = f"{Dc}{row+1}"
    return (f"=IF(A{row}>{ne}+{ns},"
            f"{S}/{Dn},"
            f"IF(A{row}>={ne},"          # >= n_ext: feed cell + all scrub use SR
            f"{S}*{R}/{Dn},"
            f"({S}*{R}+{F})/{Dn}))")

def f_d(row, elem):
    XF, F, ne = REF[f"XF_{elem}"], REF["F"], REF["n_ext"]
    return f"=IF(A{row}={ne},-{XF}*{F},0)"

def f_cp(row, bc, cc, ac, cpc):
    if row == R0:
        return f"={cc}{row}/{bc}{row}"
    w = f"({bc}{row}-{ac}{row}*{cpc}{row-1})"
    return f"={cc}{row}/{w}"

def f_dp(row, bc, dc, ac, cpc, dpc):
    if row == R0:
        return f"={dc}{row}/{bc}{row}"
    w = f"({bc}{row}-{ac}{row}*{cpc}{row-1})"
    return f"=({dc}{row}-{ac}{row}*{dpc}{row-1})/{w}"

def f_y(row, dpc, cpc, yc):
    if row == R1:
        return f"={dpc}{row}"
    return f"={dpc}{row}-{cpc}{row}*{yc}{row+1}"

# ═════════════════════════════════════════════════════════════════════════════
# BUILD WORKBOOK
# ═════════════════════════════════════════════════════════════════════════════
wb = openpyxl.Workbook()

# ─────────────────────────────────────────────────────────────────────────────
# SHEET 1: INPUTS
# ─────────────────────────────────────────────────────────────────────────────
ws = wb.active
ws.title = "INPUTS"
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 14
ws.column_dimensions["D"].width = 14
ws.column_dimensions["E"].width = 20

# Title
ws.merge_cells("A1:E1")
c = ws["A1"]
c.value = "SX STEADY-STATE MODEL  —  Alind Formulation  (D2EHPA / HCl)"
c.font = fnt(bold=True, col=WHT, sz=14)
c.fill = fill(NAVY)
c.alignment = aln()
ws.row_dimensions[1].height = 32

ws.row_dimensions[2].height = 8

# ── Section: GENERAL INPUTS ──────────────────────────────────────────────────
ws.merge_cells("A3:E3")
c = ws["A3"]
c.value = "  GENERAL INPUTS"
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(NAVY)
c.alignment = aln(h="left")
ws.row_dimensions[3].height = 20

for col, lbl in [(1,"Parameter"),(2,"Value")]:
    set_cell(ws, 4, col, value=lbl, bold=True, col_font=WHT, bg=BLUE, h="center")

gen = [
    ("Extraction Stages  (n_ext)",                  10),
    ("Scrub Stages  (n_scr)",                       10),
    ("Strip Stages  (n_str)",                        5),
    ("HCl Normality — Extraction  [N]",           0.50),
    ("HCl Normality — Scrub  [N]",                1.37),
    ("HCl Normality — Strip  [N]",                 5.0),
    ("Feed Flowrate  (F)  [L/h or gal/min]",        25),
    ("Strip Flowrate  (S)",                          12),
    ("Organic Flow  (O)",                          32.5),
    ("Reflux  (R)  — decimal  [e.g. 0.70 = 70%]", 0.70),
]
# Rows 5-14
for i, (lbl, val) in enumerate(gen, start=5):
    bg = XBLUE if i % 2 == 0 else WHT
    set_cell(ws, i, 1, value=lbl, col_font="000000", bg=bg, h="left")
    c = set_cell(ws, i, 2, value=val, col_font="000000", bg=bg)
    c.number_format = "0.00"

ws.row_dimensions[15].height = 8

# ── Section: DISTRIBUTION COEFFICIENTS ──────────────────────────────────────
ws.merge_cells("A16:E16")
c = ws["A16"]
c.value = "  DISTRIBUTION COEFFICIENTS    D = P × [HCl]ᴺ^Q"
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(NAVY)
c.alignment = aln(h="left")
ws.row_dimensions[16].height = 20

for col, lbl in [(1,"Element"),(3,"P"),(4,"Q")]:
    set_cell(ws, 17, col, value=lbl, bold=True, col_font=WHT, bg=BLUE)

dist = [
    ("Pr", 0.0008,  -1.973),
    ("Nd", 0.0031,  -2.541),
    ("Tb", 0.378,   -2.624),
    ("Dy", 0.5959,  -2.431),
]
# Rows 18-21
for i, (elem, P, Q) in enumerate(dist, start=18):
    bg = XBLUE if i % 2 == 0 else WHT
    set_cell(ws, i, 1, value=elem, col_font="000000", bg=bg)
    c = set_cell(ws, i, 3, value=P, col_font="000000", bg=bg)
    c.number_format = "0.0000"
    c = set_cell(ws, i, 4, value=Q, col_font="000000", bg=bg)
    c.number_format = "0.000"

ws.row_dimensions[22].height = 8

# ── Section: FEED CONCENTRATIONS ─────────────────────────────────────────────
ws.merge_cells("A23:E23")
c = ws["A23"]
c.value = "  FEED CONCENTRATIONS"
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(NAVY)
c.alignment = aln(h="left")
ws.row_dimensions[23].height = 20

for col, lbl in [(1,"Element"),(3,"Feed Conc. (g/L)")]:
    set_cell(ws, 24, col, value=lbl, bold=True, col_font=WHT, bg=BLUE)

feed = [("Pr", 0.0), ("Nd", 20.8), ("Tb", 0.0), ("Dy", 2.68)]
# Rows 25-28
for i, (elem, val) in enumerate(feed, start=25):
    bg = XBLUE if i % 2 == 0 else WHT
    set_cell(ws, i, 1, value=elem, col_font="000000", bg=bg)
    c = set_cell(ws, i, 3, value=val, col_font="000000", bg=bg)
    c.number_format = "0.00"

ws.row_dimensions[29].height = 8

# ── Section: COMPUTED REFERENCE VALUES ───────────────────────────────────────
ws.merge_cells("A30:E30")
c = ws["A30"]
c.value = "  COMPUTED REFERENCE VALUES  (auto-calculated)"
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(BLUE)
c.alignment = aln(h="left")
ws.row_dimensions[30].height = 20

for col, lbl in [(1,"Description"),(2,"Value")]:
    set_cell(ws, 31, col, value=lbl, bold=True, col_font=WHT, bg=BLUE)

computed = [
    ("Scrub Flowrate  =  S × R",        f"={REF['S']}*{REF['R']}"),
    ("Preg Flowrate   =  S × (1 − R)",  f"={REF['S']}*(1-{REF['R']})"),
    ("Raffinate Flowrate  =  S×R + F",  f"={REF['S']}*{REF['R']}+{REF['F']}"),
    ("Total Stages  =  n_ext+n_scr+n_str", f"={REF['n_ext']}+{REF['n_scr']}+{REF['n_str']}"),
]
for i, (lbl, frm) in enumerate(computed, start=32):
    bg = XBLUE if i % 2 == 0 else WHT
    set_cell(ws, i, 1, value=lbl, col_font="000000", bg=bg, h="left")
    c = set_cell(ws, i, 2, formula=frm, col_font="000000", bg=bg)
    c.number_format = "0.00"

# ─────────────────────────────────────────────────────────────────────────────
# SHEET 2: SOLVER
# ─────────────────────────────────────────────────────────────────────────────
ws2 = wb.create_sheet("SOLVER")

# Freeze top 4 rows + col A
ws2.freeze_panes = "B5"

# Column widths
ws2.column_dimensions["A"].width = 7   # Stage#
ws2.column_dimensions["B"].width = 10  # Section
for e in ELEMS:
    s = ESTART[e]
    for off in range(8):
        ws2.column_dimensions[gcl(s+off)].width = 11
for xc in XCOL.values():
    ws2.column_dimensions[gcl(xc)].width = 11
ws2.column_dimensions[gcl(39)].width = 12  # total org
ws2.column_dimensions[gcl(40)].width = 12  # total aq

# ── Title row 1 ───────────────────────────────────────────────────────────────
ws2.merge_cells(f"A1:{gcl(40)}1")
c = ws2["A1"]
c.value = "SOLVER  —  TDMA (Thomas Algorithm)  |  Each column-block is one REE element"
c.font = fnt(bold=True, col=WHT, sz=13)
c.fill = fill(NAVY)
c.alignment = aln()
ws2.row_dimensions[1].height = 28

# ── Row 2: element group headers ─────────────────────────────────────────────
elem_colors = {"Pr": "4472C4", "Nd": "2E75B6", "Tb": "70AD47", "Dy": ORNG}
ws2.merge_cells("A2:B2")
ws2["A2"].value = ""
ws2["A2"].fill = fill(NAVY)

for elem in ELEMS:
    s = ESTART[elem]
    ws2.merge_cells(f"{gcl(s)}2:{gcl(s+7)}2")
    c = ws2[f"{gcl(s)}2"]
    c.value = f"← {elem}  (8 columns) →"
    c.font = fnt(bold=True, col=WHT, sz=11)
    c.fill = fill(elem_colors[elem])
    c.alignment = aln()
ws2.row_dimensions[2].height = 20

# Aqueous header
ws2.merge_cells(f"{gcl(35)}2:{gcl(38)}2")
c = ws2[f"{gcl(35)}2"]
c.value = "← Aqueous  x_i  (g/L) →"
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(NAVY)
c.alignment = aln()

ws2.merge_cells(f"{gcl(39)}2:{gcl(40)}2")
c = ws2[f"{gcl(39)}2"]
c.value = "Totals"
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(NAVY)
c.alignment = aln()

# ── Row 3: sub-headers (formula meaning) ─────────────────────────────────────
sub_labels = {0:"D_i", 1:"a", 2:"b", 3:"c", 4:"d(RHS)", 5:"c′", 6:"d′", 7:"y_org"}
for elem in ELEMS:
    s = ESTART[elem]
    for off, lbl in sub_labels.items():
        c = ws2.cell(row=3, column=s+off)
        c.value = lbl
        c.font = fnt(bold=True, col=WHT, sz=9)
        c.fill = fill("555555")
        c.alignment = aln()
ws2.row_dimensions[3].height = 16

for off, lbl in enumerate(["x_Pr","x_Nd","x_Tb","x_Dy"]):
    c = ws2.cell(row=3, column=35+off)
    c.value = lbl
    c.font = fnt(bold=True, col=WHT, sz=9)
    c.fill = fill("555555")
    c.alignment = aln()

for col, lbl in [(39,"Σ y_org"),(40,"Σ x_aq")]:
    c = ws2.cell(row=3, column=col)
    c.value = lbl
    c.font = fnt(bold=True, col=WHT, sz=9)
    c.fill = fill("555555")
    c.alignment = aln()

# ── Row 4: column headers ──────────────────────────────────────────────────
set_cell(ws2, 4, 1, value="Stage", bold=True, col_font=WHT, bg=NAVY)
set_cell(ws2, 4, 2, value="Section", bold=True, col_font=WHT, bg=NAVY)
ws2.row_dimensions[4].height = 18

col4_labels = {
    0:"D = P·[HCl]^Q", 1:"a (lower diag)", 2:"b (main diag)",
    3:"c (upper diag)", 4:"d (RHS)",
    5:"c′ (fwd sweep)", 6:"d′ (fwd sweep)", 7:"y = org conc (g/L)"
}
for elem in ELEMS:
    s = ESTART[elem]
    for off, lbl in col4_labels.items():
        c = ws2.cell(row=4, column=s+off)
        c.value = lbl
        c.font = fnt(bold=True, col=WHT, sz=9)
        c.fill = fill(elem_colors[elem])
        c.alignment = aln(wrap=True)
    ws2.row_dimensions[4].height = 30

for off, lbl in enumerate(["x_Pr = y/D_Pr","x_Nd = y/D_Nd","x_Tb = y/D_Tb","x_Dy = y/D_Dy"]):
    c = ws2.cell(row=4, column=35+off)
    c.value = lbl
    c.font = fnt(bold=True, col=WHT, sz=9)
    c.fill = fill(NAVY)
    c.alignment = aln(wrap=True)

for col, lbl in [(39,"Total Organic (g/L)"),(40,"Total Aqueous (g/L)")]:
    c = ws2.cell(row=4, column=col)
    c.value = lbl
    c.font = fnt(bold=True, col=WHT, sz=9)
    c.fill = fill(NAVY)
    c.alignment = aln(wrap=True)

# ── Section boundary helpers ───────────────────────────────────────────────
def section_of(stage):
    # hard-coded for default 10/10/5
    if stage <= 10:   return "Extraction", YLW
    elif stage <= 20: return "Scrub",      LGRN
    else:             return "Strip",      LRED

# ── Data rows (stages 1-25, rows 5-29) ────────────────────────────────────
for n in range(1, N+1):
    row = R0 + n - 1  # e.g. stage 1 → row 5
    section, sec_bg = section_of(n)

    # Stage# and Section
    c = ws2.cell(row=row, column=1)
    c.value = n
    c.font = fnt(bold=True, col="000000", sz=10)
    c.fill = fill(sec_bg)
    c.alignment = aln()
    c.border = brd()

    c = ws2.cell(row=row, column=2)
    c.value = section
    c.font = fnt(col="000000", sz=9)
    c.fill = fill(sec_bg)
    c.alignment = aln()
    c.border = brd()

    for elem in ELEMS:
        s = ESTART[elem]
        Dc  = gcl(s+0); ac  = gcl(s+1); bc  = gcl(s+2)
        cc  = gcl(s+3); dc  = gcl(s+4); cpc = gcl(s+5)
        dpc = gcl(s+6); yc  = gcl(s+7)

        # bg for element (light shade of element color)
        ebg = XBLUE if n % 2 == 0 else WHT

        frm_D  = f_D(row, elem)
        frm_a  = f_a(row)
        frm_b  = f_b(row, Dc)
        frm_c  = f_c(row, Dc)
        frm_d  = f_d(row, elem)
        frm_cp = f_cp(row, bc, cc, ac, cpc)
        frm_dp = f_dp(row, bc, dc, ac, cpc, dpc)
        frm_y  = f_y(row, dpc, cpc, yc)

        for off, frm in enumerate([frm_D, frm_a, frm_b, frm_c, frm_d, frm_cp, frm_dp, frm_y]):
            c = ws2.cell(row=row, column=s+off)
            c.value = frm
            c.font = fnt(col="000000", sz=9)
            c.fill = fill(ebg)
            c.alignment = aln()
            c.border = brd()
            c.number_format = "0.0000"

    # Aqueous x_i = y_i / D_i
    for elem in ELEMS:
        s   = ESTART[elem]
        Dc  = gcl(s+0); yc = gcl(s+7)
        xcol = XCOL[elem]
        c = ws2.cell(row=row, column=xcol)
        # Protect against D=0 (shouldn't happen but guards div errors)
        c.value = f"=IF({Dc}{row}<>0,{yc}{row}/{Dc}{row},0)"
        c.font = fnt(col="000000", sz=10)
        c.fill = fill(LBLUE)
        c.alignment = aln()
        c.border = brd()
        c.number_format = "0.0000"

    # Total organic  = sum of all y_i
    c = ws2.cell(row=row, column=39)
    y_cells = "+".join(f"{gcl(ESTART[e]+7)}{row}" for e in ELEMS)
    c.value = f"={y_cells}"
    c.font = fnt(bold=True, col="000000", sz=10)
    c.fill = fill(LBLUE)
    c.alignment = aln()
    c.border = brd()
    c.number_format = "0.0000"

    # Total aqueous  = sum of all x_i
    c = ws2.cell(row=row, column=40)
    x_cells = "+".join(f"{gcl(XCOL[e])}{row}" for e in ELEMS)
    c.value = f"={x_cells}"
    c.font = fnt(bold=True, col="000000", sz=10)
    c.fill = fill(LBLUE)
    c.alignment = aln()
    c.border = brd()
    c.number_format = "0.0000"

# ─────────────────────────────────────────────────────────────────────────────
# SHEET 3: RESULTS
# ─────────────────────────────────────────────────────────────────────────────
ws3 = wb.create_sheet("RESULTS")
ws3.column_dimensions["A"].width = 24
for col in ["B","C","D","E","F","G"]:
    ws3.column_dimensions[col].width = 16

# Title
ws3.merge_cells("A1:G1")
c = ws3["A1"]
c.value = "RESULTS  —  SX Circuit Summary"
c.font = fnt(bold=True, col=WHT, sz=14)
c.fill = fill(NAVY)
c.alignment = aln()
ws3.row_dimensions[1].height = 32
ws3.row_dimensions[2].height = 8

# ── Key indicator ─────────────────────────────────────────────────────────
ws3.merge_cells("A3:G3")
c = ws3["A3"]
c.value = "  KEY METRIC"
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(NAVY)
c.alignment = aln(h="left")
ws3.row_dimensions[3].height = 20

# Row 4: Loaded Organic at extraction exit (stage n_ext = stage 10 = row 14 in SOLVER)
# Total organic at row 14 of SOLVER is column AM (col 39)
set_cell(ws3, 4, 1, value="Loaded Organic @ Ext Exit (g/L)", bold=True, col_font="000000",
         bg=YLW, h="left")
c = set_cell(ws3, 4, 2, formula="=SOLVER!AM14", col_font="000000", bg=YLW)
c.number_format = "0.00"
set_cell(ws3, 4, 3, value="[Target: 19.05]", col_font="555555", bg=YLW, border=False)

ws3.row_dimensions[5].height = 8

# ── Compositions table ────────────────────────────────────────────────────
ws3.merge_cells("A6:G6")
c = ws3["A6"]
c.value = "  COMPOSITIONS & DISTRIBUTIONS"
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(NAVY)
c.alignment = aln(h="left")
ws3.row_dimensions[6].height = 20

hdrs = ["Element", "Feed Conc.\n(g/L)", "Raff Conc.\n(g/L)", "Preg Conc.\n(g/L)",
        "% Feed", "% Raffinate", "% Preg"]
for j, lbl in enumerate(hdrs, start=1):
    c = set_cell(ws3, 7, j, value=lbl, bold=True, col_font=WHT, bg=BLUE, wrap=True)
ws3.row_dimensions[7].height = 30

# Raff = aqueous at stage 1 → SOLVER row 5 → x_i columns AI..AL (35..38)
# Preg = aqueous at first strip stage (stage n_ext+n_scr+1 = 21) → SOLVER row 25
# Row of stage n in SOLVER = R0 + n - 1 = 4 + n
raff_row = R0 + 1 - 1   # = 5  (stage 1)
preg_row = R0 + 21 - 1  # = 25 (stage 21, first strip stage)

xcol_letter = {e: gcl(XCOL[e]) for e in ELEMS}

elem_rows_in    = {"Pr": "INPUTS!$C$25", "Nd": "INPUTS!$C$26",
                   "Tb": "INPUTS!$C$27", "Dy": "INPUTS!$C$28"}

# Sum of all feed concs (denominator for % Feed)
sum_feed_frm = "+".join(f"INPUTS!$C${r}" for r in range(25, 29))
# Sum of raff concs
sum_raff_frm = "+".join(f"SOLVER!{xcol_letter[e]}{raff_row}" for e in ELEMS)
# Sum of preg concs
sum_preg_frm = "+".join(f"SOLVER!{xcol_letter[e]}{preg_row}" for e in ELEMS)

for i, elem in enumerate(ELEMS, start=8):
    bg = XBLUE if i % 2 == 0 else WHT
    xc = xcol_letter[elem]
    feed_ref = elem_rows_in[elem]

    set_cell(ws3, i, 1, value=elem, bold=True, col_font="000000", bg=bg)

    c = set_cell(ws3, i, 2, formula=f"={feed_ref}", col_font="000000", bg=bg)
    c.number_format = "0.00"

    c = set_cell(ws3, i, 3, formula=f"=SOLVER!{xc}{raff_row}", col_font="000000", bg=bg)
    c.number_format = "0.00"

    c = set_cell(ws3, i, 4, formula=f"=SOLVER!{xc}{preg_row}", col_font="000000", bg=bg)
    c.number_format = "0.00"

    # % Feed = feed_i / total_feed * 100
    c = set_cell(ws3, i, 5,
                 formula=f"=IF(({sum_feed_frm})=0,0,{feed_ref}/({sum_feed_frm})*100)",
                 col_font="000000", bg=bg)
    c.number_format = "0.00"

    # % Raffinate = x_i_raff / sum_raff * 100
    c = set_cell(ws3, i, 6,
                 formula=f"=IF(({sum_raff_frm})=0,0,"
                         f"SOLVER!{xc}{raff_row}/({sum_raff_frm})*100)",
                 col_font="000000", bg=bg)
    c.number_format = "0.00"

    # % Preg = x_i_preg / sum_preg * 100
    c = set_cell(ws3, i, 7,
                 formula=f"=IF(({sum_preg_frm})=0,0,"
                         f"SOLVER!{xc}{preg_row}/({sum_preg_frm})*100)",
                 col_font="000000", bg=bg)
    c.number_format = "0.00"

ws3.row_dimensions[13].height = 8

# ── Stage profile table (for charts) ─────────────────────────────────────
ws3.merge_cells("A14:G14")
c = ws3["A14"]
c.value = "  STAGE PROFILE  (pulled from SOLVER — used for charts)"
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(NAVY)
c.alignment = aln(h="left")
ws3.row_dimensions[14].height = 20

# Headers
for j, lbl in enumerate(["Stage","Section","x_Pr","x_Nd","x_Tb","x_Dy",
                          "x_NdPr (grouped)"], start=1):
    c = set_cell(ws3, 15, j, value=lbl, bold=True, col_font=WHT, bg=BLUE, wrap=True)
ws3.row_dimensions[15].height = 30

# Additional columns for % of total per stage (for pct chart)
for j, lbl in enumerate(["%_NdPr","%_Dy"], start=8):
    c = set_cell(ws3, 15, j, value=lbl, bold=True, col_font=WHT, bg=BLUE)

# Data rows (stages 1-25, RESULTS rows 16-40)
for n in range(1, N+1):
    row = 15 + n        # rows 16-40
    srow = R0 + n - 1   # corresponding SOLVER row

    section, sec_bg = section_of(n)
    bg = sec_bg if n % 2 != 0 else WHT

    set_cell(ws3, row, 1, value=n, col_font="000000", bg=bg)
    set_cell(ws3, row, 2, value=section, col_font="000000", bg=bg)

    for j, elem in enumerate(ELEMS, start=3):
        xc = xcol_letter[elem]
        c = set_cell(ws3, row, j,
                     formula=f"=SOLVER!{xc}{srow}",
                     col_font="000000", bg=bg)
        c.number_format = "0.0000"

    # x_NdPr = x_Nd + x_Pr (columns 4 and 3 in this table → C and D)
    c = set_cell(ws3, row, 7,
                 formula=f"=C{row}+D{row}",
                 col_font="000000", bg=LBLUE)
    c.number_format = "0.0000"

    # % NdPr = (x_Nd+x_Pr)/(total aq at this stage)*100
    # total aq at this stage = x_Pr+x_Nd+x_Tb+x_Dy = C+D+E+F
    c = set_cell(ws3, row, 8,
                 formula=f"=IF((C{row}+D{row}+E{row}+F{row})=0,0,"
                         f"(C{row}+D{row})/(C{row}+D{row}+E{row}+F{row})*100)",
                 col_font="000000", bg=XBLUE)
    c.number_format = "0.00"

    # % Dy
    c = set_cell(ws3, row, 9,
                 formula=f"=IF((C{row}+D{row}+E{row}+F{row})=0,0,"
                         f"F{row}/(C{row}+D{row}+E{row}+F{row})*100)",
                 col_font="000000", bg=XBLUE)
    c.number_format = "0.00"

# ── Chart 1: Aqueous Concentrations — Absolute (g/L) ─────────────────────
# Stage 1-25 in col A (rows 16-40), NdPr in col G, Dy in col F

chart1 = LineChart()
chart1.title       = "Aqueous Distributions — Absolute"
chart1.style       = 10
chart1.y_axis.title = "Aqueous Concentration X, g/L"
chart1.x_axis.title = "Stage"
chart1.width       = 20
chart1.height      = 14

# NdPr series (col G = 7, rows 16-40)
data_ndpr = Reference(ws3, min_col=7, max_col=7, min_row=15, max_row=40)
s1 = Series(data_ndpr, title_from_data=True)
chart1.append(s1)

# Dy series (col F = 6, rows 16-40)
data_dy = Reference(ws3, min_col=6, max_col=6, min_row=15, max_row=40)
s2 = Series(data_dy, title_from_data=True)
chart1.append(s2)

# Stage labels (col A)
cats = Reference(ws3, min_col=1, max_col=1, min_row=16, max_row=40)
chart1.set_categories(cats)

chart1.series[0].graphicalProperties.line.solidFill = "2E75B6"  # NdPr blue
chart1.series[0].graphicalProperties.line.width = 20000
chart1.series[1].graphicalProperties.line.solidFill = "ED7D31"  # Dy orange
chart1.series[1].graphicalProperties.line.width = 20000

chart1.series[0].marker.symbol = "circle"
chart1.series[1].marker.symbol = "circle"

ws3.add_chart(chart1, "A42")

# ── Chart 2: Aqueous Distributions — % of Total ──────────────────────────
chart2 = LineChart()
chart2.title        = "Aqueous Distributions — % of Total"
chart2.style        = 10
chart2.y_axis.title = "Aqueous Distribution, %"
chart2.x_axis.title = "Stage"
chart2.width        = 20
chart2.height       = 14
chart2.y_axis.numFmt = '0"%"'
chart2.y_axis.scaling.min = 0
chart2.y_axis.scaling.max = 100

# %NdPr series (col H = 8)
data_pndpr = Reference(ws3, min_col=8, max_col=8, min_row=15, max_row=40)
s3 = Series(data_pndpr, title_from_data=True)
chart2.append(s3)

# %Dy series (col I = 9)
data_pdy = Reference(ws3, min_col=9, max_col=9, min_row=15, max_row=40)
s4 = Series(data_pdy, title_from_data=True)
chart2.append(s4)

chart2.set_categories(cats)

chart2.series[0].graphicalProperties.line.solidFill = "2E75B6"
chart2.series[0].graphicalProperties.line.width = 20000
chart2.series[1].graphicalProperties.line.solidFill = "ED7D31"
chart2.series[1].graphicalProperties.line.width = 20000
chart2.series[0].marker.symbol = "circle"
chart2.series[1].marker.symbol = "circle"

ws3.add_chart(chart2, "K42")

# ── Tab colors ────────────────────────────────────────────────────────────
ws.sheet_properties.tabColor   = "1F4E79"
ws2.sheet_properties.tabColor  = "2E75B6"
ws3.sheet_properties.tabColor  = "70AD47"

# ── Save ──────────────────────────────────────────────────────────────────
out = "/home/user/SX-simulation/SX_Steady_State_Model.xlsx"
wb.save(out)
print(f"Saved: {out}")
