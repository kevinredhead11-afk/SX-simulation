# SX Steady-State Model — Step-by-Step Excel Guide

**System:** D2EHPA / HCl  |  **Reference:** Alind Chandra formulation  
**Solver:** TDMA (Thomas Algorithm) implemented as direct Excel formulas  
**Elements modelled:** Pr, Nd, Tb, Dy  

---

## What This File Does

This Excel workbook simulates a continuous solvent extraction (SX) circuit at **steady state**.
Given a set of operating conditions, it calculates what concentration of each rare earth
element (REE) will be found in every mixer-settler stage of the circuit, and where each
element ends up — in the raffinate or in the pregnant leach solution (preg).

The workbook has three sheets:

| Sheet | Purpose |
|---|---|
| **INPUTS** | All parameters you can change |
| **SOLVER** | Matrix construction + TDMA solution (do not edit) |
| **RESULTS** | Summary table and two charts |

---

## Part 1 — Chemistry Background

### The Reaction

When rare earth chlorides (LnCl₃) in aqueous HCl contact D2EHPA (abbreviated H‑DEHP) dissolved in kerosene, the following exchange reaction occurs:

```
3 H-DEHP(org)  +  LnCl₃(aq)  ⇌  Ln-DEHP₃(org)  +  3 HCl(aq)
```

The REE transfers from aqueous to organic; acid is released back into the aqueous phase.

### The Distribution Coefficient

The distribution coefficient **D** is defined as:

```
D  =  [Ln in organic phase]  /  [Ln in aqueous phase]
```

For a fixed D2EHPA concentration and temperature, D depends only on the HCl normality:

```
D = P × [HCl]^Q
```

Where **P** and **Q** are element-specific empirical constants determined by shakeout tests.
A higher D means the element prefers the organic phase at that acid concentration.

Key insight: **Q is always negative**, so increasing HCl normality drives REEs *back* into
the aqueous phase. This is why:
- **Extraction** uses low HCl (~0.5 N) → high D → REEs load onto organic
- **Scrub** uses moderate HCl (~1.37 N) → lower D → only the most strongly loaded REEs stay
- **Strip** uses high HCl (~5 N) → very low D → REEs are forced back to aqueous (preg)

---

## Part 2 — Circuit Topology

The circuit has three sections wired in series. Organic flows in one direction; aqueous
flows countercurrent. Here is how the stages are numbered in this model:

```
Stage:     1  2  3 ... 10 | 11 12 ... 20 | 21 22 ... 25
Section:   ←— Extraction —→ ←——— Scrub ———→ ←—— Strip ——→
           (n_ext = 10)      (n_scr = 10)    (n_str = 5)
```

**Organic flow direction →** (stage 1 to 25, then recycles back)  
**Aqueous flow direction ←** (countercurrent in each section)

Key streams:

| Stream | Location | Flow rate |
|---|---|---|
| Feed | Aqueous enters at stage **n_ext** (stage 10) | F |
| Raffinate | Aqueous leaves at stage **1** | S×R + F |
| Scrub acid | Aqueous enters at stage **n_ext + 1** (stage 11) — as reflux from strip | S×R |
| Preg | Aqueous leaves at stage **n_ext + n_scr + 1** (stage 21) | S×(1−R) |
| Strip acid | Aqueous enters at last stage **25** | S |
| Loaded organic | Organic leaves extraction at stage 10 | O |

**Reflux R** is the fraction of the strip section aqueous outflow that is recycled back
into the scrub section as scrub solution (instead of exiting as preg).

---

## Part 3 — Mathematical Model

### Setting Up the Mass Balances

For each stage, the amount of REE entering equals the amount leaving (steady state).
Both organic (y) and aqueous (x) streams carry REE, and they are linked by equilibrium:

```
y_{i,n}  =  k_{i,n} × x_{i,n}
```

where k = D (the distribution coefficient at the HCl normality of that section).

After substituting the equilibrium relation, each stage n gives one equation in terms
of the **organic concentrations y only**:

```
a_n × y_{n-1}  +  b_n × y_n  +  c_n × y_{n+1}  =  d_n
```

This is a **tridiagonal equation** — each unknown (y_n) is connected only to its immediate
neighbors. The full system for all N stages forms a tridiagonal matrix.

### The a, b, c, d Coefficients

The coefficients depend on which section the stage belongs to:

**Strip section** (stages n_ext+n_scr+1 to N):
```
a_n  =  O                    (organic from previous stage)
b_n  =  −O − S / D_n         (total outflow from stage n)
c_n  =  S / D_{n+1}          (aqueous contribution from next stage)
d_n  =  0
```
Exception: last stage (n = N): `c_N = 0` (no stage beyond N), and d contains strip acid
  concentration, but since strip acid carries no REE, d_N = 0.

**Scrub section** (stages n_ext+1 to n_ext+n_scr):
```
a_n  =  O
b_n  =  −O − (S×R) / D_n
c_n  =  (S×R) / D_{n+1}
d_n  =  0
```

**Extraction section** (stages 1 to n_ext):
```
a_n  =  O    (except a_1 = 0, as no stage precedes stage 1)
b_n  =  −O − (S×R + F) / D_n
c_n  =  (S×R + F) / D_{n+1}    for n < n_ext
c_n  =  (S×R) / D_{n+1}         for n = n_ext (feed cell; F enters here, not carried forward)
d_n  =  0    except at the feed cell (n = n_ext):
d_{n_ext}  =  −(Feed_concentration) × F
```

The rightmost column `d` is zero everywhere except the feed stage, which acts as the
source term injecting the feed into the system.

---

## Part 4 — Solving with TDMA (Thomas Algorithm)

The tridiagonal system cannot be solved in one step with a direct formula, but it can be
solved in two linear sweeps — **no iteration required**.

### Step 1 — Forward Elimination (top to bottom)

Compute modified coefficients c′ and d′:

```
c′_1  =  c_1 / b_1
d′_1  =  d_1 / b_1

For n = 2 to N:
  w_n   =  b_n  −  a_n × c′_{n-1}
  c′_n  =  c_n / w_n
  d′_n  =  (d_n  −  a_n × d′_{n-1}) / w_n
```

In the SOLVER sheet, these are the **c′ and d′ columns** for each element. Each row
references only the row above — pure top-down calculation.

### Step 2 — Back Substitution (bottom to top)

```
y_N  =  d′_N

For n = N−1 down to 1:
  y_n  =  d′_n  −  c′_n × y_{n+1}
```

In the SOLVER sheet, the **y_org column** for each element references the row *below* it.
Excel resolves this automatically through its dependency engine — no circular references.

### Getting Aqueous Concentrations

Once y_n (organic) is known, the aqueous concentration follows directly:

```
x_n  =  y_n / D_n
```

These are the **x_Pr, x_Nd, x_Tb, x_Dy columns** (AI–AL) in the SOLVER sheet.

---

## Part 5 — The INPUTS Sheet

Open the INPUTS sheet and change only the **yellow/white cells** in the value columns.
Never edit the formula cells in SOLVER or RESULTS.

### General Inputs

| Cell | Parameter | Description |
|---|---|---|
| B5 | n_ext | Number of extraction stages |
| B6 | n_scr | Number of scrub stages |
| B7 | n_str | Number of strip stages |
| B8 | HCl Ext [N] | HCl normality in extraction section |
| B9 | HCl Scr [N] | HCl normality in scrub section |
| B10 | HCl Str [N] | HCl normality in strip section |
| B11 | F | Feed flowrate (any consistent units) |
| B12 | S | Strip acid flowrate |
| B13 | O | Organic flowrate |
| B14 | R | Reflux fraction (0.70 = 70%) |

> **Important:** If you change n_ext, n_scr, or n_str, the SOLVER sheet rows are fixed at 25 stages.
> You must keep n_ext + n_scr + n_str = 25.

### Distribution Coefficients (D = P × [HCl]^Q)

| Rows | Columns | Data |
|---|---|---|
| 18–21 | C | P constant for Pr, Nd, Tb, Dy |
| 18–21 | D | Q exponent for Pr, Nd, Tb, Dy |

These constants come from shakeout experiments. P sets the overall extractability;
Q (always negative) sets how strongly acid suppresses extraction.

### Feed Concentrations

| Rows | Column | Data |
|---|---|---|
| 25–28 | C | Feed concentration of Pr, Nd, Tb, Dy in g/L |

---

## Part 6 — The SOLVER Sheet

**Do not edit any cell in this sheet.** All cells contain formulas.

### Column Organization

Each element (Pr, Nd, Tb, Dy) occupies 8 consecutive columns:

```
Offset 0: D_i      = P × [HCl]^Q  for the section of this stage
Offset 1: a        = sub-diagonal coefficient
Offset 2: b        = main-diagonal coefficient
Offset 3: c        = super-diagonal coefficient
Offset 4: d        = right-hand side (only non-zero at feed stage)
Offset 5: c′       = forward-sweep modified super-diagonal
Offset 6: d′       = forward-sweep modified RHS
Offset 7: y_org    = organic concentration (g/L)  ← SOLUTION
```

Column blocks:
- **Pr**: columns C–J
- **Nd**: columns K–R
- **Tb**: columns S–Z
- **Dy**: columns AA–AH

After these blocks:
- **AI–AL**: aqueous concentrations x_Pr, x_Nd, x_Tb, x_Dy = y/D
- **AM**: total organic g/L (sum of all y)
- **AN**: total aqueous g/L (sum of all x)

### Row Colors

| Color | Section |
|---|---|
| Yellow | Extraction stages (1–10) |
| Green | Scrub stages (11–20) |
| Red/orange | Strip stages (21–25) |

### How to Read the Results

- **Stage 14 (row 14), column AM**: Loaded organic at extraction exit = total REE on organic entering scrub
- **Stage 1 (row 5), columns AI–AL**: Raffinate concentrations
- **Stage 21 (row 25), columns AI–AL**: Preg concentrations

---

## Part 7 — The RESULTS Sheet

### Key Metric (Row 4)

**Loaded Organic @ Ext Exit** is the total REE loading on the organic leaving the last
extraction stage (entering scrub). This is a primary circuit design indicator.
Target value for the reference case: **19.05 g/L**.

### Compositions & Distributions Table (Rows 8–11)

| Column | Meaning |
|---|---|
| Feed Conc. | Input concentration from INPUTS sheet |
| Raff Conc. | Aqueous concentration at stage 1 (raffinate) |
| Preg Conc. | Aqueous concentration at stage 21 (first strip stage = preg exit) |
| % Feed | Fraction of total feed mass that is this element |
| % Raffinate | Purity of this element in the raffinate stream |
| % Preg | Purity of this element in the preg stream |

### Stage Profile Table (Rows 16–40)

Aqueous concentrations at every stage, pulled from SOLVER. Columns:
- x_Pr, x_Nd, x_Tb, x_Dy — individual element concentrations
- x_NdPr (grouped) — Nd + Pr combined (matches screenshot display)
- %_NdPr, %_Dy — fraction of total aqueous REE at each stage

### Charts

**Chart 1 — Aqueous Distributions, Absolute**: shows g/L of NdPr and Dy at each stage.
Matches the lower-right chart in the reference screenshot.

**Chart 2 — Aqueous Distributions, % of Total**: shows the composition as a percentage
of total REE in the aqueous phase at each stage. Matches the upper-right chart.

Both charts use stages 1–25 on the X-axis. The crossover point between stages 10 and 11
(extraction/scrub boundary) is where separation occurs.

---

## Part 8 — Validation Against Reference Results

With the default inputs, the model should reproduce:

| Output | Expected Value |
|---|---|
| Loaded Organic at stage 10 | 19.05 g/L |
| Raffinate Nd conc. | ~15.57 g/L |
| Raffinate purity | 100% Nd |
| Preg Dy conc. | ~18.61 g/L |
| Preg purity | 100% Dy |
| % Feed — Nd | 88.59% |
| % Feed — Dy | 11.41% |

If values differ, check:
1. All INPUTS values exactly match the reference parameters
2. The SOLVER sheet has no `#DIV/0!` or `#VALUE!` errors (check columns b and c′)
3. Excel calculation mode is set to **Automatic** (Formulas → Calculation Options → Automatic)

---

## Part 9 — How to Use for Design Studies

### Changing Operating Conditions

1. Open INPUTS sheet
2. Modify any value in column B (General Inputs) or column C (Feed Concentrations)
3. Results update automatically — no button click required

### Understanding Sensitivity

| Parameter | Increasing it... |
|---|---|
| O (organic flow) | Loads more REE onto organic, may overwhelm scrub |
| S (strip flow) | Improves stripping but dilutes preg; check Preg Conc. |
| R (reflux) | Higher R → cleaner separation, lower preg flow |
| HCl Ext | Higher acid → harder to load organic (lower D) |
| HCl Scr | Higher acid → more aggressive scrub, pushes weakly-loaded REEs to raff |
| n_ext | More stages → more complete loading (up to equilibrium limit) |
| n_scr | More stages → higher purity in preg |

### Changing the Separation Target

To separate a different pair of elements, update the P and Q values in INPUTS for those
elements. The model will automatically recalculate. The key is that the two elements you
want to separate must have different D values at the operating HCl normalities — expressed
as a **separation factor β = D_heavy / D_light**.

---

## Part 10 — Excel Formula Reference

Below is a translation of the key Excel formulas used in the SOLVER sheet, so you can
verify or rebuild them if needed.

### D coefficient (cell C5, Pr at stage 1)
```
=IF(A5<=INPUTS!$B$5,
    INPUTS!$C$18 * INPUTS!$B$8 ^ INPUTS!$D$18,
    IF(A5<=INPUTS!$B$5+INPUTS!$B$6,
       INPUTS!$C$18 * INPUTS!$B$9 ^ INPUTS!$D$18,
       INPUTS!$C$18 * INPUTS!$B$10 ^ INPUTS!$D$18))
```
Logic: if stage ≤ n_ext → use Ext HCl; elif ≤ n_ext+n_scr → use Scr HCl; else Strip HCl.

### b coefficient (cell E5, Pr at stage 1)
```
=IF(A5>INPUTS!$B$5+INPUTS!$B$6,
    -INPUTS!$B$13 - INPUTS!$B$12/C5,
    IF(A5>INPUTS!$B$5,
       -INPUTS!$B$13 - INPUTS!$B$12*INPUTS!$B$14/C5,
       -INPUTS!$B$13 - (INPUTS!$B$12*INPUTS!$B$14+INPUTS!$B$11)/C5))
```

### c′ forward sweep (cell H5, stage 1 — base case)
```
=F5/E5
```
For stage n > 1:
```
=(F{n}) / (E{n} - D{n}*H{n-1})
```

### y back substitution (cell J29, stage 25 — last stage)
```
=I29
```
For stage n < 25:
```
=I{n} - H{n}*J{n+1}
```

### Aqueous concentration (cell AI5, x_Pr at stage 1)
```
=IF(C5<>0, J5/C5, 0)
```

---

*Guide version 1.0 — Alind SX Steady-State Model, D2EHPA/HCl, TDMA formulation*
