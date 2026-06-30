#!/usr/bin/env python3
"""
SX_Steady_State_Model.xlsx  —  SINGLE SHEET VERSION
All inputs, solver, results, and charts on one sheet.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as gcl

# ── Layout constants ──────────────────────────────────────────────────────────
N   = 25          # fixed total stages
R0  = 20          # first SOLVER data row (stage 1)
R1  = R0 + N - 1  # last SOLVER data row  (stage 25 = row 44)

# SOLVER column map (1-indexed, same as before)
ELEMS  = ["Pr", "Nd", "Tb", "Dy"]
ESTART = {"Pr": 3, "Nd": 11, "Tb": 19, "Dy": 27}  # D col of each element block
XCOL   = {"Pr": 35, "Nd": 36, "Tb": 37, "Dy": 38}  # x_i aqueous cols

# Stage profile table (for charts) starts at:
PROF_HDR = R1 + 3          # = 47  header row
PROF_R0  = PROF_HDR + 1    # = 48  first data row
PROF_R1  = PROF_R0 + N - 1 # = 72  last data row

# ── INPUTS cell refs (all in same sheet, no sheet prefix) ─────────────────────
# General inputs: col B, rows 5-14
# Dist coefficients: col H (P) and I (Q), rows 5-8
# Feed concentrations: col M, rows 5-8
REF = {
    "n_ext": "$B$5",  "n_scr": "$B$6",  "n_str": "$B$7",
    "He":    "$B$8",  "Hs":    "$B$9",  "Ht":    "$B$10",
    "F":     "$B$11", "S":     "$B$12",
    "O":     "$B$13", "R":     "$B$14",
    "P_Pr":  "$H$5",  "Q_Pr":  "$I$5",  "XF_Pr": "$M$5",
    "P_Nd":  "$H$6",  "Q_Nd":  "$I$6",  "XF_Nd": "$M$6",
    "P_Tb":  "$H$7",  "Q_Tb":  "$I$7",  "XF_Tb": "$M$7",
    "P_Dy":  "$H$8",  "Q_Dy":  "$I$8",  "XF_Dy": "$M$8",
}

# ── Colours ───────────────────────────────────────────────────────────────────
NAVY  = "1F4E79"; BLUE  = "2E75B6"; LBLUE = "BDD7EE"; XBLUE = "EBF3FA"
ORNG  = "ED7D31"; GRN   = "70AD47"; RED   = "C00000"
YLW   = "FFF2CC"; LGRN  = "E2EFDA"; LRED  = "FCE4D6"
WHT   = "FFFFFF"; GRY   = "D9D9D9"; DGRY  = "555555"

ELEM_CLR = {"Pr": "4472C4", "Nd": "2E75B6", "Tb": "70AD47", "Dy": "ED7D31"}

# ── Style helpers ─────────────────────────────────────────────────────────────
def fill(c): return PatternFill("solid", fgColor=c)
def fnt(bold=False, col="000000", sz=11):
    return Font(bold=bold, color=col, size=sz, name="Calibri")
def aln(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
_t = Side(style="thin",   color="BBBBBB")
_m = Side(style="medium", color="1F4E79")
def brd():  return Border(left=_t, right=_t, top=_t, bottom=_t)
def brdm(): return Border(left=_m, right=_m, top=_m, bottom=_m)

def sc(ws, row, col, val=None, frm=None, bold=False, fc=WHT, sz=11,
       bg=None, h="center", wrap=False, nf=None, bdr=True):
    c = ws.cell(row=row, column=col)
    c.value = frm if frm is not None else val
    c.font = fnt(bold=bold, col=fc, sz=sz)
    if bg: c.fill = fill(bg)
    c.alignment = aln(h=h, wrap=wrap)
    if bdr: c.border = brd()
    if nf:  c.number_format = nf
    return c

def hdr(ws, row, col1, col2, text, bg=NAVY, sz=11):
    ws.merge_cells(f"{gcl(col1)}{row}:{gcl(col2)}{row}")
    c = ws.cell(row=row, column=col1)
    c.value = text
    c.font = fnt(bold=True, col=WHT, sz=sz)
    c.fill = fill(bg)
    c.alignment = aln(h="left")
    ws.row_dimensions[row].height = 20

# ── TDMA formula builders ─────────────────────────────────────────────────────
def f_D(row, elem):
    P, Q = REF[f"P_{elem}"], REF[f"Q_{elem}"]
    ne, ns = REF["n_ext"], REF["n_scr"]
    He, Hs, Ht = REF["He"], REF["Hs"], REF["Ht"]
    return (f"=IF(A{row}<={ne},{P}*{He}^{Q},"
            f"IF(A{row}<={ne}+{ns},{P}*{Hs}^{Q},{P}*{Ht}^{Q}))")

def f_a(row):
    return f"=IF(A{row}=1,0,{REF['O']})"

def f_b(row, Dc):
    O,S,R,F = REF["O"],REF["S"],REF["R"],REF["F"]
    ne,ns,D = REF["n_ext"],REF["n_scr"],f"{Dc}{row}"
    return (f"=IF(A{row}>{ne}+{ns},-{O}-{S}/{D},"
            f"IF(A{row}>{ne},-{O}-{S}*{R}/{D},-{O}-({S}*{R}+{F})/{D}))")

def f_c(row, Dc):
    S,R,F = REF["S"],REF["R"],REF["F"]
    ne,ns = REF["n_ext"],REF["n_scr"]
    if row == R1: return "=0"
    Dn = f"{Dc}{row+1}"
    return (f"=IF(A{row}>{ne}+{ns},{S}/{Dn},"
            f"IF(A{row}>={ne},{S}*{R}/{Dn},({S}*{R}+{F})/{Dn}))")

def f_d(row, elem):
    XF,F,ne = REF[f"XF_{elem}"],REF["F"],REF["n_ext"]
    return f"=IF(A{row}={ne},-{XF}*{F},0)"

def f_cp(row, bc, cc, ac, cpc):
    if row == R0: return f"={cc}{row}/{bc}{row}"
    w = f"({bc}{row}-{ac}{row}*{cpc}{row-1})"
    return f"={cc}{row}/{w}"

def f_dp(row, bc, dc, ac, cpc, dpc):
    if row == R0: return f"={dc}{row}/{bc}{row}"
    w = f"({bc}{row}-{ac}{row}*{cpc}{row-1})"
    return f"=({dc}{row}-{ac}{row}*{dpc}{row-1})/{w}"

def f_y(row, dpc, cpc, yc):
    if row == R1: return f"={dpc}{row}"
    return f"={dpc}{row}-{cpc}{row}*{yc}{row+1}"

# ═════════════════════════════════════════════════════════════════════════════
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "SX Model"
ws.sheet_view.showGridLines = True

# ── Column widths ─────────────────────────────────────────────────────────────
ws.column_dimensions["A"].width = 30   # parameter labels / stage#
ws.column_dimensions["B"].width = 12   # values
ws.column_dimensions["C"].width = 2    # spacer
ws.column_dimensions["D"].width = 2    # spacer
ws.column_dimensions["E"].width = 2    # spacer
ws.column_dimensions["F"].width = 10   # elem label
ws.column_dimensions["G"].width = 2    # spacer
ws.column_dimensions["H"].width = 12   # P
ws.column_dimensions["I"].width = 12   # Q
ws.column_dimensions["J"].width = 2    # spacer
ws.column_dimensions["K"].width = 10   # elem label (feed)
ws.column_dimensions["L"].width = 2    # spacer
ws.column_dimensions["M"].width = 14   # feed conc
ws.column_dimensions["N"].width = 2    # spacer
ws.column_dimensions["O"].width = 36   # results label
ws.column_dimensions["P"].width = 14   # results value

for e in ELEMS:
    s = ESTART[e]
    for off in range(8):
        ws.column_dimensions[gcl(s+off)].width = 11
for xc in XCOL.values():
    ws.column_dimensions[gcl(xc)].width = 11
ws.column_dimensions[gcl(39)].width = 12
ws.column_dimensions[gcl(40)].width = 12

# ═════════════════════════════════════════════════════════════════════════════
# ROW 1 — MAIN TITLE
# ═════════════════════════════════════════════════════════════════════════════
ws.merge_cells(f"A1:{gcl(40)}1")
c = ws["A1"]
c.value = "SX STEADY-STATE MODEL  —  Alind Formulation  (D2EHPA / HCl)  —  TDMA Solver"
c.font = fnt(bold=True, col=WHT, sz=14)
c.fill = fill(NAVY)
c.alignment = aln()
ws.row_dimensions[1].height = 34
ws.row_dimensions[2].height = 8

# ═════════════════════════════════════════════════════════════════════════════
# ROWS 3-14 — INPUTS SECTION (three panels side by side)
# ═════════════════════════════════════════════════════════════════════════════

# ── Panel 1: General Inputs (cols A-B, rows 3-14) ────────────────────────────
hdr(ws, 3, 1, 2, "  GENERAL INPUTS")
sc(ws, 4, 1, val="Parameter",        bold=True, fc=WHT, bg=BLUE, h="left")
sc(ws, 4, 2, val="Value",            bold=True, fc=WHT, bg=BLUE)

gen = [
    ("Extraction Stages  (n_ext)",                  10),
    ("Scrub Stages  (n_scr)",                       10),
    ("Strip Stages  (n_str)",                        5),
    ("HCl Normality — Extraction  [N]",           0.50),
    ("HCl Normality — Scrub  [N]",                1.37),
    ("HCl Normality — Strip  [N]",                 5.0),
    ("Feed Flowrate  (F)",                           25),
    ("Strip Flowrate  (S)",                          12),
    ("Organic Flow  (O)",                          32.5),
    ("Reflux  (R)  [decimal: 0.70 = 70%]",        0.70),
]
for i, (lbl, val) in enumerate(gen, start=5):
    bg = XBLUE if i % 2 == 0 else WHT
    sc(ws, i, 1, val=lbl, fc="000000", bg=bg, h="left")
    sc(ws, i, 2, val=val, fc="000000", bg=bg, nf="0.00")
# rows 5-14: B5=n_ext … B14=R

# ── Panel 2: Distribution Coefficients (cols F-I, rows 3-9) ─────────────────
hdr(ws, 3, 6, 9, "  DIST. COEFFICIENTS    D = P × [HCl]^Q")
for col, lbl in [(6,"Element"),(8,"P"),(9,"Q")]:
    sc(ws, 4, col, val=lbl, bold=True, fc=WHT, bg=BLUE)
dist = [("Pr",0.0008,-1.973),("Nd",0.0031,-2.541),("Tb",0.378,-2.624),("Dy",0.5959,-2.431)]
for i,(elem,P,Q) in enumerate(dist, start=5):
    bg = XBLUE if i%2==0 else WHT
    sc(ws,i,6,val=elem, fc="000000",bg=bg)
    sc(ws,i,8,val=P,    fc="000000",bg=bg,nf="0.0000")
    sc(ws,i,9,val=Q,    fc="000000",bg=bg,nf="0.000")
# H5=P_Pr,I5=Q_Pr … H8=P_Dy,I8=Q_Dy

# ── Panel 3: Feed Concentrations (cols K-M, rows 3-9) ────────────────────────
hdr(ws, 3, 11, 13, "  FEED CONCENTRATIONS")
for col,lbl in [(11,"Element"),(13,"Feed Conc. (g/L)")]:
    sc(ws,4,col,val=lbl,bold=True,fc=WHT,bg=BLUE)
feed = [("Pr",0.0),("Nd",20.8),("Tb",0.0),("Dy",2.68)]
for i,(elem,val) in enumerate(feed,start=5):
    bg = XBLUE if i%2==0 else WHT
    sc(ws,i,11,val=elem, fc="000000",bg=bg)
    sc(ws,i,13,val=val,  fc="000000",bg=bg,nf="0.00")
# M5=Feed_Pr, M6=Feed_Nd, M7=Feed_Tb, M8=Feed_Dy

# ── Panel 4: Key Results Summary (cols O-P, rows 3-14) ───────────────────────
hdr(ws, 3, 15, 16, "  RESULTS SUMMARY  (auto-updated)")
sc(ws,4,15,val="Metric",       bold=True,fc=WHT,bg=BLUE,h="left")
sc(ws,4,16,val="Value",        bold=True,fc=WHT,bg=BLUE)

# These reference SOLVER data using INDEX — fully dynamic
ne, ns = REF["n_ext"], REF["n_scr"]
# Stage n is at SOLVER row R0 + n - 1  =  19 + n
# loaded organic: stage n_ext → row 19+n_ext
lo_row  = f"19+{ne}"
# preg: first strip stage = n_ext+n_scr+1 → row 19+n_ext+n_scr+1 = 20+n_ext+n_scr
preg_idx = f"20+{ne}+{ns}"
raff_row = R0   # stage 1 → fixed row

xcl = {e: gcl(XCOL[e]) for e in ELEMS}

def idx(col_ltr, row_expr): return f"INDEX(${col_ltr}:${col_ltr},{row_expr})"
def raff(e):  return f"${xcl[e]}${raff_row}"
def preg(e):  return idx(xcl[e], preg_idx)
def xtot(cells): return "+".join(cells)

sum_feed = "+".join(REF[f"XF_{e}"] for e in ELEMS)
sum_raff = xtot([raff(e) for e in ELEMS])
sum_preg = xtot([preg(e) for e in ELEMS])

results_rows = [
    ("Loaded Organic @ Ext Exit (g/L)",
     f"=INDEX($AM:$AM,{lo_row})", "0.00", YLW),
    ("Raffinate — Nd (g/L)",
     f"={raff('Nd')}", "0.00", XBLUE),
    ("Raffinate — Dy (g/L)",
     f"={raff('Dy')}", "0.00", XBLUE),
    ("Preg — Nd (g/L)",
     f"={preg('Nd')}", "0.00", LGRN),
    ("Preg — Dy (g/L)",
     f"={preg('Dy')}", "0.00", LGRN),
    ("% Feed  —  Nd",
     f"=IF(({sum_feed})=0,0,{REF['XF_Nd']}/({sum_feed})*100)", "0.00", WHT),
    ("% Feed  —  Dy",
     f"=IF(({sum_feed})=0,0,{REF['XF_Dy']}/({sum_feed})*100)", "0.00", WHT),
    ("% Purity in Raffinate  —  Nd",
     f"=IF(({sum_raff})=0,0,{raff('Nd')}/({sum_raff})*100)", "0.00", WHT),
    ("% Purity in Preg  —  Dy",
     f"=IF(({sum_preg})=0,0,{preg('Dy')}/({sum_preg})*100)", "0.00", WHT),
    ("Scrub Flowrate  =  S × R",
     f"={REF['S']}*{REF['R']}", "0.00", GRY),
    ("Preg Flowrate   =  S × (1−R)",
     f"={REF['S']}*(1-{REF['R']})", "0.00", GRY),
]
for i,(lbl,frm,nf,bg) in enumerate(results_rows, start=5):
    sc(ws,i,15,val=lbl, fc="000000",bg=bg,h="left")
    sc(ws,i,16,frm=frm, fc="000000",bg=bg,nf=nf)

# ═════════════════════════════════════════════════════════════════════════════
# ROWS 16-19 — SOLVER SECTION HEADERS
# ═════════════════════════════════════════════════════════════════════════════
ws.row_dimensions[15].height = 8

ws.merge_cells(f"A16:{gcl(40)}16")
c = ws["A16"]
c.value = ("SOLVER  —  TDMA (Thomas Algorithm)  |  Each 8-column block = one REE element"
           "  |  Forward sweep (c′,d′) top→bottom  |  Back-sub (y_org) bottom→top")
c.font = fnt(bold=True, col=WHT, sz=11)
c.fill = fill(NAVY)
c.alignment = aln(h="left")
ws.row_dimensions[16].height = 22

# Row 17: element group banners
ws.merge_cells("A17:B17")
ws["A17"].fill = fill(NAVY)
for e in ELEMS:
    s = ESTART[e]
    ws.merge_cells(f"{gcl(s)}17:{gcl(s+7)}17")
    c = ws[f"{gcl(s)}17"]
    c.value = f"← {e}  (8 cols) →"
    c.font = fnt(bold=True, col=WHT, sz=10)
    c.fill = fill(ELEM_CLR[e])
    c.alignment = aln()
ws.merge_cells(f"{gcl(35)}17:{gcl(38)}17")
c = ws[f"{gcl(35)}17"]; c.value="← Aqueous x_i →"; c.font=fnt(bold=True,col=WHT,sz=10); c.fill=fill(NAVY); c.alignment=aln()
ws.merge_cells(f"{gcl(39)}17:{gcl(40)}17")
c = ws[f"{gcl(39)}17"]; c.value="Totals"; c.font=fnt(bold=True,col=WHT,sz=10); c.fill=fill(NAVY); c.alignment=aln()
ws.row_dimensions[17].height = 18

# Row 18: sub-labels
sublbls = {0:"D_i",1:"a",2:"b",3:"c",4:"d (RHS)",5:"c′",6:"d′",7:"y_org"}
for e in ELEMS:
    s = ESTART[e]
    for off,lbl in sublbls.items():
        c = ws.cell(row=18, column=s+off)
        c.value=lbl; c.font=fnt(bold=True,col=WHT,sz=8); c.fill=fill(DGRY); c.alignment=aln()
for off,lbl in enumerate(["x_Pr","x_Nd","x_Tb","x_Dy"]):
    c=ws.cell(row=18,column=35+off); c.value=lbl; c.font=fnt(bold=True,col=WHT,sz=8); c.fill=fill(DGRY); c.alignment=aln()
for col,lbl in [(39,"Σ y_org"),(40,"Σ x_aq")]:
    c=ws.cell(row=18,column=col); c.value=lbl; c.font=fnt(bold=True,col=WHT,sz=8); c.fill=fill(DGRY); c.alignment=aln()
ws.row_dimensions[18].height = 14

# Row 19: full column header descriptions
sc(ws,19,1,val="Stage",   bold=True,fc=WHT,bg=NAVY,sz=9)
sc(ws,19,2,val="Section", bold=True,fc=WHT,bg=NAVY,sz=9)
col19 = {0:"D=P·[HCl]^Q",1:"a (sub-diag)",2:"b (main diag)",
         3:"c (super-diag)",4:"d (RHS)",5:"c′ fwd sweep",6:"d′ fwd sweep",7:"y org (g/L)"}
for e in ELEMS:
    s = ESTART[e]
    for off,lbl in col19.items():
        c=ws.cell(row=19,column=s+off); c.value=lbl
        c.font=fnt(bold=True,col=WHT,sz=8); c.fill=fill(ELEM_CLR[e]); c.alignment=aln(wrap=True)
for off,lbl in enumerate(["x_Pr=y/D","x_Nd=y/D","x_Tb=y/D","x_Dy=y/D"]):
    c=ws.cell(row=19,column=35+off); c.value=lbl; c.font=fnt(bold=True,col=WHT,sz=8); c.fill=fill(NAVY); c.alignment=aln(wrap=True)
for col,lbl in [(39,"Total Org (g/L)"),(40,"Total Aq (g/L)")]:
    c=ws.cell(row=19,column=col); c.value=lbl; c.font=fnt(bold=True,col=WHT,sz=8); c.fill=fill(NAVY); c.alignment=aln(wrap=True)
ws.row_dimensions[19].height = 28

# ═════════════════════════════════════════════════════════════════════════════
# ROWS 20-44 — SOLVER DATA (stages 1-25)
# ═════════════════════════════════════════════════════════════════════════════
for n in range(1, N+1):
    row = R0 + n - 1   # row 20..44

    # Stage# — color via conditional formatting applied below
    c=ws.cell(row=row,column=1); c.value=n; c.font=fnt(bold=True,col="000000",sz=10)
    c.fill=fill(WHT); c.alignment=aln(); c.border=brd()

    # Section label — dynamic formula
    c=ws.cell(row=row,column=2)
    c.value=(f'=IF(A{row}<={REF["n_ext"]},"Extraction",'
             f'IF(A{row}<={REF["n_ext"]}+{REF["n_scr"]},"Scrub","Strip"))')
    c.font=fnt(col="000000",sz=9); c.fill=fill(WHT); c.alignment=aln(); c.border=brd()

    for e in ELEMS:
        s   = ESTART[e]
        Dc  = gcl(s);   ac  = gcl(s+1); bc  = gcl(s+2)
        cc  = gcl(s+3); dc  = gcl(s+4); cpc = gcl(s+5)
        dpc = gcl(s+6); yc  = gcl(s+7)
        ebg = XBLUE if n%2==0 else WHT

        formulas = [f_D(row,e), f_a(row), f_b(row,Dc), f_c(row,Dc), f_d(row,e),
                    f_cp(row,bc,cc,ac,cpc), f_dp(row,bc,dc,ac,cpc,dpc), f_y(row,dpc,cpc,yc)]
        for off,frm in enumerate(formulas):
            c=ws.cell(row=row,column=s+off); c.value=frm
            c.font=fnt(col="000000",sz=9); c.fill=fill(ebg)
            c.alignment=aln(); c.border=brd(); c.number_format="0.0000"

    # Aqueous x_i = y/D
    for e in ELEMS:
        s=ESTART[e]; Dc=gcl(s); yc=gcl(s+7); xc=XCOL[e]
        c=ws.cell(row=row,column=xc)
        c.value=f"=IF({Dc}{row}<>0,{yc}{row}/{Dc}{row},0)"
        c.font=fnt(col="000000",sz=10); c.fill=fill(LBLUE)
        c.alignment=aln(); c.border=brd(); c.number_format="0.0000"

    # Total organic
    c=ws.cell(row=row,column=39)
    c.value="="+"+".join(f"{gcl(ESTART[e]+7)}{row}" for e in ELEMS)
    c.font=fnt(bold=True,col="000000",sz=10); c.fill=fill(LBLUE)
    c.alignment=aln(); c.border=brd(); c.number_format="0.0000"

    # Total aqueous
    c=ws.cell(row=row,column=40)
    c.value="="+"+".join(f"{gcl(XCOL[e])}{row}" for e in ELEMS)
    c.font=fnt(bold=True,col="000000",sz=10); c.fill=fill(LBLUE)
    c.alignment=aln(); c.border=brd(); c.number_format="0.0000"

# ── Conditional formatting for section colors (cols A:B, rows 20-44) ─────────
cf_rng = f"A{R0}:B{R1}"
ws.conditional_formatting.add(cf_rng,
    FormulaRule(formula=[f"$A{R0}>{REF['n_ext']}+{REF['n_scr']}"],
                fill=PatternFill("solid",fgColor=LRED),  stopIfTrue=False))
ws.conditional_formatting.add(cf_rng,
    FormulaRule(formula=[f"$A{R0}>{REF['n_ext']}"],
                fill=PatternFill("solid",fgColor=LGRN),  stopIfTrue=False))
ws.conditional_formatting.add(cf_rng,
    FormulaRule(formula=[f"$A{R0}<={REF['n_ext']}"],
                fill=PatternFill("solid",fgColor=YLW),   stopIfTrue=False))

# ═════════════════════════════════════════════════════════════════════════════
# STAGE PROFILE TABLE (for charts)
# ═════════════════════════════════════════════════════════════════════════════
ws.row_dimensions[R1+2].height = 8

ws.merge_cells(f"A{PROF_HDR}:{gcl(9)}{PROF_HDR}")
c=ws.cell(row=PROF_HDR,column=1)
c.value="STAGE PROFILE  (aqueous concentrations, g/L)  —  used for charts below"
c.font=fnt(bold=True,col=WHT,sz=11); c.fill=fill(NAVY); c.alignment=aln(h="left")
ws.row_dimensions[PROF_HDR].height=20

for col,lbl in enumerate(["Stage","Section","x_Pr","x_Nd","x_Tb","x_Dy",
                           "x_NdPr\n(grouped)","%_NdPr","%_Dy"],start=1):
    sc(ws,PROF_HDR+1,col,val=lbl,bold=True,fc=WHT,bg=BLUE,wrap=True)
ws.row_dimensions[PROF_HDR+1].height=30

def sec_of(n):
    if n<=10: return "Extraction",YLW
    elif n<=20: return "Scrub",LGRN
    else: return "Strip",LRED

for n in range(1,N+1):
    pr = PROF_R0 + n - 1   # profile row
    sr = R0 + n - 1        # solver row
    sec,bg = sec_of(n)

    sc(ws,pr,1,val=n,  fc="000000",bg=bg)
    sc(ws,pr,2,val=sec,fc="000000",bg=bg)

    for j,e in enumerate(ELEMS,start=3):
        xc=gcl(XCOL[e])
        c=sc(ws,pr,j,frm=f"=${xc}${sr}",fc="000000",bg=bg); c.number_format="0.0000"

    # x_NdPr = x_Nd + x_Pr (cols D+C in profile = cols 4+3)
    c=sc(ws,pr,7,frm=f"=C{pr}+D{pr}",fc="000000",bg=LBLUE); c.number_format="0.0000"

    tot = f"(C{pr}+D{pr}+E{pr}+F{pr})"
    c=sc(ws,pr,8,frm=f"=IF({tot}=0,0,(C{pr}+D{pr})/{tot}*100)",
         fc="000000",bg=XBLUE); c.number_format="0.00"
    c=sc(ws,pr,9,frm=f"=IF({tot}=0,0,F{pr}/{tot}*100)",
         fc="000000",bg=XBLUE); c.number_format="0.00"

# ═════════════════════════════════════════════════════════════════════════════
# CHARTS
# ═════════════════════════════════════════════════════════════════════════════
cats = Reference(ws, min_col=1, max_col=1, min_row=PROF_R0, max_row=PROF_R1)

# ── Chart 1: Absolute concentrations ─────────────────────────────────────────
ch1 = LineChart()
ch1.title        = "Aqueous Distributions — Absolute"
ch1.style        = 10
ch1.width        = 22
ch1.height       = 14
ch1.x_axis.title = "Stage"
ch1.y_axis.title = "Aqueous Concentration (g/L)"
ch1.x_axis.numFmt = "0"
ch1.y_axis.numFmt = "0.00"
ch1.x_axis.tickLblPos = "low"
ch1.legend.position   = "tr"

s_ndpr = Series(Reference(ws,min_col=7,max_col=7,min_row=PROF_HDR+1,max_row=PROF_R1),
                title_from_data=True)
s_dy   = Series(Reference(ws,min_col=6,max_col=6,min_row=PROF_HDR+1,max_row=PROF_R1),
                title_from_data=True)
ch1.append(s_ndpr); ch1.append(s_dy)
ch1.set_categories(cats)

ch1.series[0].graphicalProperties.line.solidFill = "2E75B6"
ch1.series[0].graphicalProperties.line.width     = 22000
ch1.series[0].marker.symbol  = "circle"
ch1.series[0].marker.size    = 5
ch1.series[1].graphicalProperties.line.solidFill = "ED7D31"
ch1.series[1].graphicalProperties.line.width     = 22000
ch1.series[1].marker.symbol  = "circle"
ch1.series[1].marker.size    = 5

ws.add_chart(ch1, f"A{PROF_R1+3}")

# ── Chart 2: % of Total ───────────────────────────────────────────────────────
ch2 = LineChart()
ch2.title        = "Aqueous Distributions — % of Total"
ch2.style        = 10
ch2.width        = 22
ch2.height       = 14
ch2.x_axis.title = "Stage"
ch2.y_axis.title = "Aqueous Distribution (%)"
ch2.x_axis.numFmt = "0"
ch2.y_axis.numFmt = "0"
ch2.x_axis.tickLblPos = "low"
ch2.legend.position   = "tr"
ch2.y_axis.scaling.min = 0
ch2.y_axis.scaling.max = 100

s_pndpr = Series(Reference(ws,min_col=8,max_col=8,min_row=PROF_HDR+1,max_row=PROF_R1),
                 title_from_data=True)
s_pdy   = Series(Reference(ws,min_col=9,max_col=9,min_row=PROF_HDR+1,max_row=PROF_R1),
                 title_from_data=True)
ch2.append(s_pndpr); ch2.append(s_pdy)
ch2.set_categories(cats)

ch2.series[0].graphicalProperties.line.solidFill = "2E75B6"
ch2.series[0].graphicalProperties.line.width     = 22000
ch2.series[0].marker.symbol  = "circle"
ch2.series[0].marker.size    = 5
ch2.series[1].graphicalProperties.line.solidFill = "ED7D31"
ch2.series[1].graphicalProperties.line.width     = 22000
ch2.series[1].marker.symbol  = "circle"
ch2.series[1].marker.size    = 5

ws.add_chart(ch2, f"M{PROF_R1+3}")

# ── Tab color and freeze panes ────────────────────────────────────────────────
ws.sheet_properties.tabColor = "1F4E79"
ws.freeze_panes = "C20"   # freeze rows 1-19 and col A-B while scrolling solver

# ── Save ──────────────────────────────────────────────────────────────────────
out = "/home/user/SX-simulation/SX_Steady_State_Model.xlsx"
wb.save(out)
print(f"Saved: {out}")
print(f"INPUTS:  rows 3-14, cols A-P")
print(f"SOLVER:  rows 16-44, cols A-AN ({gcl(40)})")
print(f"PROFILE: rows {PROF_HDR}-{PROF_R1}, cols A-I")
print(f"CHARTS:  rows {PROF_R1+3}+")
