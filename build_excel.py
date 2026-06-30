#!/usr/bin/env python3
"""
SX_Steady_State_Model.xlsx  —  Professional single-sheet version
Color system:
  INPUT  = pale yellow  #FFFDE7 / navy font    — hard-coded values the user edits
  CALC   = white        #FFFFFF  / black font   — intermediate formula cells
  STEP   = pale blue    #E8EAF6 / black font    — TDMA forward-sweep intermediate
  OUTPUT = pale green   #F1F8E9 / dark-green    — solution & final results
  REF    = pale lavender #EDE7F6 / black        — cells that pull from elsewhere
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as gcl

# ── Layout constants ──────────────────────────────────────────────────────────
N       = 25
R0      = 20           # first SOLVER data row (stage 1)
R1      = R0 + N - 1  # last  SOLVER data row (stage 25 = row 44)
PROF_HDR = R1 + 3      # stage-profile header row = 47
PROF_R0  = PROF_HDR + 1
PROF_R1  = PROF_R0 + N - 1

ELEMS  = ["Pr", "Nd", "Tb", "Dy"]
ESTART = {"Pr": 3, "Nd": 11, "Tb": 19, "Dy": 27}
XCOL   = {"Pr": 35, "Nd": 36, "Tb": 37, "Dy": 38}

# ── Cell refs (same sheet, no prefix) ────────────────────────────────────────
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

# ── Professional colour system ────────────────────────────────────────────────
# Structure colours
T_TITLE  = "1B3A6B"   # deep navy     — main title
T_SECHDR = "2E4057"   # dark slate    — section headers
T_COLHDR = "4A6FA5"   # steel blue    — column headers
T_LEGEND = "37474F"   # blue-grey     — legend bar

# Semantic cell colours
I_BG = "FFFDE7"  # pale yellow  — INPUT  (hardcoded values)
I_FG = "1A237E"  # dark navy    — INPUT  font
C_BG = "FFFFFF"  # white        — CALC   (formula, intermediate)
C_FG = "212121"  # near-black   — CALC   font
S_BG = "E8EAF6"  # pale indigo  — STEP   (TDMA c′ d′ — intermediate sweep)
S_FG = "212121"  # near-black   — STEP   font
O_BG = "F1F8E9"  # pale mint    — OUTPUT (y_org solution, final results)
O_FG = "2E7D32"  # dark green   — OUTPUT font
R_BG = "E3F2FD"  # pale sky     — REF    (x_aq derived, profile pulls)
R_FG = "0D47A1"  # dark blue    — REF    font
WHT  = "FFFFFF"

# Solver element header colours (muted, professional)
E_CLR = {"Pr": "546E7A",   # blue-grey
          "Nd": "455A64",   # darker blue-grey
          "Tb": "37474F",   # darkest blue-grey
          "Dy": "4E6073"}   # slate

# Stage section tints (very subtle — section identity only)
SEC_EXT = "FFFDE7"   # same as INPUT bg — extraction = loading
SEC_SCR = "F3E5F5"   # pale lavender    — scrub
SEC_STR = "E8F5E9"   # pale mint        — strip = output

# ── Style helpers ─────────────────────────────────────────────────────────────
def fill(c): return PatternFill("solid", fgColor=c)

def font(bold=False, col="212121", sz=10):
    return Font(bold=bold, color=col, size=sz, name="Calibri")

def aln(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def side(w="thin",  col="D0D0D0"): return Side(style=w, color=col)
def side_m(col="9E9E9E"):           return Side(style="medium", color=col)

def brd(col="D0D0D0"):
    s = side(col=col)
    return Border(left=s, right=s, top=s, bottom=s)

def brd_section():
    m = side_m(); t = side()
    return Border(left=m, right=m, top=t, bottom=t)

def style_input(c, val=None, frm=None, nf=None, h="center"):
    c.value = frm if frm is not None else val
    c.font  = font(col=I_FG)
    c.fill  = fill(I_BG)
    c.alignment = aln(h=h)
    c.border = brd()
    if nf: c.number_format = nf

def style_calc(c, frm, nf=None, h="center", sz=9):
    c.value = frm
    c.font  = font(col=C_FG, sz=sz)
    c.fill  = fill(C_BG)
    c.alignment = aln(h=h)
    c.border = brd()
    if nf: c.number_format = nf

def style_step(c, frm, nf=None):
    c.value = frm
    c.font  = font(col=S_FG, sz=9)
    c.fill  = fill(S_BG)
    c.alignment = aln()
    c.border = brd()
    if nf: c.number_format = nf

def style_output(c, frm=None, val=None, nf=None, h="center", sz=10):
    c.value = frm if frm is not None else val
    c.font  = font(bold=True, col=O_FG, sz=sz)
    c.fill  = fill(O_BG)
    c.alignment = aln(h=h)
    c.border = brd()
    if nf: c.number_format = nf

def style_ref(c, frm, nf=None, h="center"):
    c.value = frm
    c.font  = font(col=R_FG, sz=9)
    c.fill  = fill(R_BG)
    c.alignment = aln(h=h)
    c.border = brd()
    if nf: c.number_format = nf

def hdr_main(ws, row, col1, col2, text, bg=T_TITLE, sz=11):
    ws.merge_cells(f"{gcl(col1)}{row}:{gcl(col2)}{row}")
    c = ws.cell(row=row, column=col1)
    c.value = text
    c.font  = font(bold=True, col=WHT, sz=sz)
    c.fill  = fill(bg)
    c.alignment = aln(h="left")
    return c

def hdr_col(ws, row, col, text, bg=T_COLHDR, sz=9, wrap=False):
    c = ws.cell(row=row, column=col)
    c.value = text
    c.font  = font(bold=True, col=WHT, sz=sz)
    c.fill  = fill(bg)
    c.alignment = aln(wrap=wrap)
    c.border = brd("AAAAAA")
    return c

def label(ws, row, col, text, bg=None, fc="212121", bold=False, h="left", sz=10):
    c = ws.cell(row=row, column=col)
    c.value = text
    c.font  = font(bold=bold, col=fc, sz=sz)
    if bg: c.fill = fill(bg)
    c.alignment = aln(h=h)
    c.border = brd()
    return c

# ── TDMA formula builders ─────────────────────────────────────────────────────
def f_D(row, elem):
    P,Q   = REF[f"P_{elem}"], REF[f"Q_{elem}"]
    ne,ns = REF["n_ext"], REF["n_scr"]
    He,Hs,Ht = REF["He"], REF["Hs"], REF["Ht"]
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

# ── Column widths ─────────────────────────────────────────────────────────────
col_widths = {
    "A":30, "B":12, "C":2,  "D":2,  "E":2,
    "F":10, "G":2,  "H":12, "I":12,
    "J":2,  "K":10, "L":2,  "M":14,
    "N":2,  "O":36, "P":14
}
for c,w in col_widths.items():
    ws.column_dimensions[c].width = w
for e in ELEMS:
    for off in range(8):
        ws.column_dimensions[gcl(ESTART[e]+off)].width = 11
for xc in XCOL.values():
    ws.column_dimensions[gcl(xc)].width = 11
ws.column_dimensions[gcl(39)].width = 12
ws.column_dimensions[gcl(40)].width = 12

# ═════════════════════════════════════════════════════════════════════════════
# ROW 1 — MAIN TITLE
# ═════════════════════════════════════════════════════════════════════════════
ws.merge_cells(f"A1:{gcl(40)}1")
c = ws["A1"]
c.value = "SX STEADY-STATE MODEL  ·  Alind Formulation  ·  D2EHPA / HCl  ·  TDMA Direct Solver"
c.font  = font(bold=True, col=WHT, sz=13)
c.fill  = fill(T_TITLE)
c.alignment = aln()
ws.row_dimensions[1].height = 30

# ROW 2 — colour legend bar
ws.row_dimensions[2].height = 18
legend = [
    (1,  4,  "INPUT — hard-coded value",    I_BG, I_FG),
    (5,  8,  "CALC — intermediate formula", C_BG, C_FG),
    (9,  12, "STEP — TDMA sweep (c′ d′)",   S_BG, S_FG),
    (13, 16, "OUTPUT — solution / result",  O_BG, O_FG),
    (17, 20, "REF — derived / linked",      R_BG, R_FG),
]
for c1, c2, text, bg, fg in legend:
    ws.merge_cells(f"{gcl(c1)}2:{gcl(c2)}2")
    c = ws.cell(row=2, column=c1)
    c.value = text
    c.font  = font(bold=True, col=fg, sz=9)
    c.fill  = fill(bg)
    c.alignment = aln()
    c.border = brd("BBBBBB")

# ═════════════════════════════════════════════════════════════════════════════
# ROWS 4-15 — INPUTS  (three panels + results summary)
# ═════════════════════════════════════════════════════════════════════════════
ws.row_dimensions[3].height = 6

# ── Panel A: General Inputs — cols A-B, rows 4-15 ────────────────────────────
hdr_main(ws, 4, 1, 2, "  GENERAL INPUTS", bg=T_SECHDR, sz=10)
hdr_col(ws, 5, 1, "Parameter", bg=T_COLHDR)
hdr_col(ws, 5, 2, "Value",     bg=T_COLHDR)

gen = [
    ("Extraction Stages  (n_ext)",               10),
    ("Scrub Stages  (n_scr)",                    10),
    ("Strip Stages  (n_str)",                     5),
    ("HCl Normality — Extraction  [N]",         0.50),
    ("HCl Normality — Scrub  [N]",              1.37),
    ("HCl Normality — Strip  [N]",               5.0),
    ("Feed Flowrate  (F)",                         25),
    ("Strip Flowrate  (S)",                        12),
    ("Organic Flow  (O)",                        32.5),
    ("Reflux  (R)  [decimal: 0.70 = 70%]",      0.70),
]
for i, (lbl, val) in enumerate(gen, start=6):
    label(ws, i, 1, lbl, bg=I_BG, fc=I_FG)
    c = ws.cell(row=i, column=2)
    style_input(c, val=val, nf="0.00")
ws.row_dimensions[5].height = 16
# B6=n_ext, B7=n_scr, B8=n_str, B9=He, B10=Hs, B11=Ht, B12=F, B13=S, B14=O, B15=R

# ── Panel B: Distribution Coefficients — cols F-I, rows 4-9 ─────────────────
hdr_main(ws, 4, 6, 9, "  DISTRIBUTION COEFFICIENTS   D = P × [HCl]^Q", bg=T_SECHDR, sz=10)
for col, lbl in [(6,"Element"),(8,"P"),(9,"Q")]:
    hdr_col(ws, 5, col, lbl, bg=T_COLHDR)
dist = [("Pr",0.0008,-1.973),("Nd",0.0031,-2.541),("Tb",0.378,-2.624),("Dy",0.5959,-2.431)]
for i,(elem,P,Q) in enumerate(dist, start=6):
    label(ws, i, 6, elem, bg=I_BG, fc=I_FG, h="center")
    c=ws.cell(row=i,column=8); style_input(c, val=P, nf="0.0000")
    c=ws.cell(row=i,column=9); style_input(c, val=Q, nf="0.000")
# H6=P_Pr,I6=Q_Pr … H9=P_Dy,I9=Q_Dy

# ── Panel C: Feed Concentrations — cols K-M, rows 4-9 ────────────────────────
hdr_main(ws, 4, 11, 13, "  FEED CONCENTRATIONS", bg=T_SECHDR, sz=10)
for col,lbl in [(11,"Element"),(13,"Feed (g/L)")]:
    hdr_col(ws, 5, col, lbl, bg=T_COLHDR)
feed = [("Pr",0.0),("Nd",20.8),("Tb",0.0),("Dy",2.68)]
for i,(elem,val) in enumerate(feed, start=6):
    label(ws, i, 11, elem, bg=I_BG, fc=I_FG, h="center")
    c=ws.cell(row=i,column=13); style_input(c, val=val, nf="0.00")
# M6=XF_Pr, M7=XF_Nd, M8=XF_Tb, M9=XF_Dy

# ── Panel D: Results Summary — cols O-P, rows 4-15 ───────────────────────────
hdr_main(ws, 4, 15, 16, "  RESULTS SUMMARY", bg=T_SECHDR, sz=10)
hdr_col(ws, 5, 15, "Metric",  bg=T_COLHDR)
hdr_col(ws, 5, 16, "Value",   bg=T_COLHDR)

# Dynamic INDEX references — adjust offsets for new row layout
# gen inputs now start at row 6 (not 5): n_ext=$B$6, n_scr=$B$7, etc.
# Update REF to match new row layout
REF.update({
    "n_ext": "$B$6",  "n_scr": "$B$7",  "n_str": "$B$8",
    "He":    "$B$9",  "Hs":    "$B$10", "Ht":    "$B$11",
    "F":     "$B$12", "S":     "$B$13",
    "O":     "$B$14", "R":     "$B$15",
    "P_Pr":  "$H$6",  "Q_Pr":  "$I$6",  "XF_Pr": "$M$6",
    "P_Nd":  "$H$7",  "Q_Nd":  "$I$7",  "XF_Nd": "$M$7",
    "P_Tb":  "$H$8",  "Q_Tb":  "$I$8",  "XF_Tb": "$M$8",
    "P_Dy":  "$H$9",  "Q_Dy":  "$I$9",  "XF_Dy": "$M$9",
})

ne, ns = REF["n_ext"], REF["n_scr"]
lo_row   = f"19+{ne}"       # SOLVER row of stage n_ext  (R0-1+n = 19+n)
preg_idx = f"20+{ne}+{ns}"  # SOLVER row of first strip stage
raff_row = R0               # stage 1 always at R0 = row 20

xcl = {e: gcl(XCOL[e]) for e in ELEMS}
def idx(col_ltr, row_expr): return f"INDEX(${col_ltr}:${col_ltr},{row_expr})"
def raff(e):  return f"${xcl[e]}${raff_row}"
def preg(e):  return idx(xcl[e], preg_idx)
sum_feed = "+".join(REF[f"XF_{e}"] for e in ELEMS)
sum_raff = "+".join(raff(e) for e in ELEMS)
sum_preg = "+".join(preg(e) for e in ELEMS)

results = [
    ("Loaded Organic @ Extraction Exit (g/L)", f"=INDEX($AM:$AM,{lo_row})","0.00"),
    ("Raffinate — Nd  (g/L)",   f"={raff('Nd')}", "0.00"),
    ("Raffinate — Dy  (g/L)",   f"={raff('Dy')}", "0.00"),
    ("Preg — Nd  (g/L)",        f"={preg('Nd')}", "0.00"),
    ("Preg — Dy  (g/L)",        f"={preg('Dy')}", "0.00"),
    ("% Feed  —  Nd", f"=IF(({sum_feed})=0,0,{REF['XF_Nd']}/({sum_feed})*100)","0.00"),
    ("% Feed  —  Dy", f"=IF(({sum_feed})=0,0,{REF['XF_Dy']}/({sum_feed})*100)","0.00"),
    ("% Purity in Raffinate  —  Nd",
     f"=IF(({sum_raff})=0,0,{raff('Nd')}/({sum_raff})*100)","0.00"),
    ("% Purity in Preg  —  Dy",
     f"=IF(({sum_preg})=0,0,{preg('Dy')}/({sum_preg})*100)","0.00"),
    ("Scrub Flowrate  =  S × R",     f"={REF['S']}*{REF['R']}","0.00"),
    ("Preg Flowrate   =  S × (1−R)", f"={REF['S']}*(1-{REF['R']})","0.00"),
]
for i,(lbl,frm,nf) in enumerate(results, start=6):
    label(ws, i, 15, lbl, bg=O_BG, fc=O_FG)
    c=ws.cell(row=i,column=16); style_output(c, frm=frm, nf=nf, sz=10)

# ═════════════════════════════════════════════════════════════════════════════
# SOLVER HEADER ROWS 17-19
# ═════════════════════════════════════════════════════════════════════════════
ws.row_dimensions[16].height = 6

# Row 17: SOLVER banner
ws.merge_cells(f"A17:{gcl(40)}17")
c = ws["A17"]
c.value = ("SOLVER  ·  TDMA (Thomas Algorithm)  ·  "
           "Columns:  D_i  |  a  b  c  d  (matrix)  |  c′  d′  (forward sweep →)  |  "
           "y_org  (back substitution ←)  |  x_aq  (aqueous output)")
c.font  = font(bold=True, col=WHT, sz=10)
c.fill  = fill(T_SECHDR)
c.alignment = aln(h="left")
ws.row_dimensions[17].height = 20

# Row 18: element group banners
ws.merge_cells("A18:B18")
ws["A18"].fill = fill(T_SECHDR)
for e in ELEMS:
    s = ESTART[e]
    ws.merge_cells(f"{gcl(s)}18:{gcl(s+7)}18")
    c = ws[f"{gcl(s)}18"]
    c.value = f"{e}  (8 columns)"
    c.font  = font(bold=True, col=WHT, sz=10)
    c.fill  = fill(E_CLR[e])
    c.alignment = aln()
ws.merge_cells(f"{gcl(35)}18:{gcl(38)}18")
c=ws[f"{gcl(35)}18"]; c.value="Aqueous  x_i"; c.font=font(bold=True,col=WHT,sz=10); c.fill=fill(T_SECHDR); c.alignment=aln()
ws.merge_cells(f"{gcl(39)}18:{gcl(40)}18")
c=ws[f"{gcl(39)}18"]; c.value="Totals"; c.font=font(bold=True,col=WHT,sz=10); c.fill=fill(T_SECHDR); c.alignment=aln()
ws.row_dimensions[18].height = 18

# Row 19: column headers with colour showing cell type
sub = {0:("D_i",C_BG),1:("a",C_BG),2:("b",C_BG),3:("c",C_BG),4:("d (RHS)",C_BG),
       5:("c′ →",S_BG),6:("d′ →",S_BG),7:("y_org ←",O_BG)}
hdr_col(ws, 19, 1, "Stage",   bg=T_COLHDR, sz=9)
hdr_col(ws, 19, 2, "Section", bg=T_COLHDR, sz=9)
for e in ELEMS:
    s = ESTART[e]
    for off,(lbl,cbg) in sub.items():
        c = ws.cell(row=19, column=s+off)
        c.value = lbl
        c.font  = font(bold=True, col=T_TITLE if cbg==C_BG else (O_FG if cbg==O_BG else S_FG), sz=8)
        c.fill  = fill(cbg)
        c.alignment = aln(wrap=True)
        c.border = brd("AAAAAA")
for off,lbl in enumerate(["x_Pr","x_Nd","x_Tb","x_Dy"]):
    c=ws.cell(row=19,column=35+off); c.value=lbl
    c.font=font(bold=True,col=R_FG,sz=8); c.fill=fill(R_BG); c.alignment=aln(); c.border=brd("AAAAAA")
for col,lbl in [(39,"Σ Organic"),(40,"Σ Aqueous")]:
    c=ws.cell(row=19,column=col); c.value=lbl
    c.font=font(bold=True,col=O_FG,sz=8); c.fill=fill(O_BG); c.alignment=aln(); c.border=brd("AAAAAA")
ws.row_dimensions[19].height = 26

# ═════════════════════════════════════════════════════════════════════════════
# SOLVER DATA ROWS 20-44
# ═════════════════════════════════════════════════════════════════════════════
def sec_bg(n):
    if n<=10: return SEC_EXT
    elif n<=20: return SEC_SCR
    else: return SEC_STR

for n in range(1, N+1):
    row = R0 + n - 1
    sbg = sec_bg(n)

    # Stage# and Section label
    c=ws.cell(row=row,column=1); c.value=n
    c.font=font(bold=True,col=T_TITLE,sz=10); c.fill=fill(sbg); c.alignment=aln(); c.border=brd()
    c=ws.cell(row=row,column=2)
    c.value=(f'=IF(A{row}<={REF["n_ext"]},"Extraction",'
             f'IF(A{row}<={REF["n_ext"]}+{REF["n_scr"]},"Scrub","Strip"))')
    c.font=font(col=T_TITLE,sz=9); c.fill=fill(sbg); c.alignment=aln(); c.border=brd()

    for e in ELEMS:
        s   = ESTART[e]
        Dc  = gcl(s);   ac  = gcl(s+1); bc  = gcl(s+2)
        cc  = gcl(s+3); dc  = gcl(s+4); cpc = gcl(s+5)
        dpc = gcl(s+6); yc  = gcl(s+7)

        frms = [f_D(row,e), f_a(row), f_b(row,Dc), f_c(row,Dc), f_d(row,e),
                f_cp(row,bc,cc,ac,cpc), f_dp(row,bc,dc,ac,cpc,dpc), f_y(row,dpc,cpc,yc)]
        styles = [C_BG,C_BG,C_BG,C_BG,C_BG, S_BG,S_BG, O_BG]
        fgs    = [C_FG,C_FG,C_FG,C_FG,C_FG, S_FG,S_FG, O_FG]
        bolds  = [False]*5 + [False,False,True]

        for off,(frm,bg,fg,bd) in enumerate(zip(frms,styles,fgs,bolds)):
            c = ws.cell(row=row, column=s+off)
            c.value = frm
            c.font  = font(bold=bd, col=fg, sz=9)
            c.fill  = fill(bg)
            c.alignment = aln()
            c.border = brd()
            c.number_format = "0.0000"

    # x_i = y/D  (REF style)
    for e in ELEMS:
        s=ESTART[e]; Dc=gcl(s); yc=gcl(s+7)
        c=ws.cell(row=row,column=XCOL[e])
        c.value=f"=IF({Dc}{row}<>0,{yc}{row}/{Dc}{row},0)"
        c.font=font(col=R_FG,sz=9); c.fill=fill(R_BG)
        c.alignment=aln(); c.border=brd(); c.number_format="0.0000"

    # Total organic (OUTPUT)
    c=ws.cell(row=row,column=39)
    c.value="="+"+".join(f"{gcl(ESTART[e]+7)}{row}" for e in ELEMS)
    c.font=font(bold=True,col=O_FG,sz=9); c.fill=fill(O_BG)
    c.alignment=aln(); c.border=brd(); c.number_format="0.0000"

    # Total aqueous (OUTPUT)
    c=ws.cell(row=row,column=40)
    c.value="="+"+".join(f"{gcl(XCOL[e])}{row}" for e in ELEMS)
    c.font=font(bold=True,col=O_FG,sz=9); c.fill=fill(O_BG)
    c.alignment=aln(); c.border=brd(); c.number_format="0.0000"

# Conditional formatting — section tints on cols A:B
cf = f"A{R0}:B{R1}"
ws.conditional_formatting.add(cf, FormulaRule(
    formula=[f"$A{R0}>{REF['n_ext']}+{REF['n_scr']}"],
    fill=PatternFill("solid",fgColor=SEC_STR), stopIfTrue=False))
ws.conditional_formatting.add(cf, FormulaRule(
    formula=[f"$A{R0}>{REF['n_ext']}"],
    fill=PatternFill("solid",fgColor=SEC_SCR), stopIfTrue=False))
ws.conditional_formatting.add(cf, FormulaRule(
    formula=[f"$A{R0}<={REF['n_ext']}"],
    fill=PatternFill("solid",fgColor=SEC_EXT), stopIfTrue=False))

# ═════════════════════════════════════════════════════════════════════════════
# STAGE PROFILE TABLE (for charts)
# ═════════════════════════════════════════════════════════════════════════════
ws.row_dimensions[R1+2].height = 6

hdr_main(ws, PROF_HDR, 1, 9,
         "  STAGE PROFILE  ·  Aqueous concentrations (g/L)  ·  Source for charts",
         bg=T_SECHDR, sz=10)
ws.row_dimensions[PROF_HDR].height = 20

for col,lbl in enumerate(["Stage","Section","x_Pr","x_Nd","x_Tb","x_Dy",
                           "x_NdPr\n(Nd+Pr)","%_NdPr","%_Dy"],start=1):
    hdr_col(ws, PROF_HDR+1, col, lbl, bg=T_COLHDR, sz=9, wrap=True)
ws.row_dimensions[PROF_HDR+1].height = 28

def sec_label(n):
    if n<=10: return "Extraction", SEC_EXT
    elif n<=20: return "Scrub",    SEC_SCR
    else: return "Strip",          SEC_STR

for n in range(1,N+1):
    pr = PROF_R0 + n - 1
    sr = R0      + n - 1
    sec, sbg = sec_label(n)

    c=ws.cell(row=pr,column=1); c.value=n
    c.font=font(col=T_TITLE,sz=9); c.fill=fill(sbg); c.alignment=aln(); c.border=brd()
    c=ws.cell(row=pr,column=2); c.value=sec
    c.font=font(col=T_TITLE,sz=9); c.fill=fill(sbg); c.alignment=aln(); c.border=brd()

    for j,e in enumerate(ELEMS,start=3):
        c=ws.cell(row=pr,column=j)
        style_ref(c, frm=f"=${gcl(XCOL[e])}${sr}", nf="0.0000")

    # x_NdPr
    c=ws.cell(row=pr,column=7)
    style_ref(c, frm=f"=C{pr}+D{pr}", nf="0.0000")

    tot=f"(C{pr}+D{pr}+E{pr}+F{pr})"
    c=ws.cell(row=pr,column=8)
    style_output(c, frm=f"=IF({tot}=0,0,(C{pr}+D{pr})/{tot}*100)", nf="0.00", sz=9)
    c=ws.cell(row=pr,column=9)
    style_output(c, frm=f"=IF({tot}=0,0,F{pr}/{tot}*100)", nf="0.00", sz=9)

# ═════════════════════════════════════════════════════════════════════════════
# CHARTS
# ═════════════════════════════════════════════════════════════════════════════
cats = Reference(ws, min_col=1, max_col=1, min_row=PROF_R0, max_row=PROF_R1)

def make_chart(title, y_title, col1, col2, y_min=None, y_max=None, pct_fmt=False):
    ch = LineChart()
    ch.title  = title
    ch.style  = 2        # clean minimal style
    ch.width  = 22
    ch.height = 14
    ch.x_axis.title    = "Stage"
    ch.y_axis.title    = y_title
    ch.x_axis.numFmt   = "0"
    ch.y_axis.numFmt   = '0"%"' if pct_fmt else "0.00"
    ch.x_axis.tickLblPos = "low"
    ch.legend.position = "tr"
    if y_min is not None: ch.y_axis.scaling.min = y_min
    if y_max is not None: ch.y_axis.scaling.max = y_max

    for col, color, label_text in [(col1,"1B3A6B",None),(col2,"C07B30",None)]:
        s = Series(Reference(ws, min_col=col, max_col=col,
                             min_row=PROF_HDR+1, max_row=PROF_R1),
                   title_from_data=True)
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width     = 22000
        s.marker.symbol = "circle"
        s.marker.size   = 5
        s.marker.graphicalProperties.solidFill   = color
        s.marker.graphicalProperties.line.solidFill = color
        ch.append(s)

    ch.set_categories(cats)
    return ch

ch1 = make_chart("Aqueous Distributions — Absolute",
                 "Aqueous Concentration  X  (g/L)", col1=7, col2=6)
ch2 = make_chart("Aqueous Distributions — % of Total",
                 "Aqueous Distribution  (%)",
                 col1=8, col2=9, y_min=0, y_max=100, pct_fmt=True)

ws.add_chart(ch1, f"A{PROF_R1+3}")
ws.add_chart(ch2, f"M{PROF_R1+3}")

# ── Tab colour ────────────────────────────────────────────────────────────────
ws.sheet_properties.tabColor = "1B3A6B"

# ── Save ──────────────────────────────────────────────────────────────────────
out = "/home/user/SX-simulation/SX_Steady_State_Model.xlsx"
wb.save(out)
print(f"Saved: {out}")
print(f"Layout:  legend row 2 | inputs rows 4-15 | solver rows 17-44 | profile rows {PROF_HDR}-{PROF_R1} | charts rows {PROF_R1+3}+")
