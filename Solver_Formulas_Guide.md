# SOLVER Sheet — Every Formula Explained Step by Step

**What the SOLVER sheet does:**
For each of the 25 stages and each of the 4 elements (Pr, Nd, Tb, Dy), it builds a
system of equations that describes where each element goes at steady state, then solves
that system directly using the Thomas Algorithm (TDMA). No iteration, no guessing —
one clean pass down, one pass up, and you have the answer.

Each element occupies **8 columns** in the SOLVER sheet. The columns appear in this order:

```
D_i  │  a  │  b  │  c  │  d  │  c′  │  d′  │  y_org
```

Then after all four elements, there are two more columns:

```
x_i (aqueous)  │  Σ y_org  │  Σ x_aq
```

Below is a complete explanation of every formula, in the order they are evaluated.

---

## Column 1 — Stage Number (Column A)

```
A5 = 1
A6 = 2
...
A29 = 25
```

Just the stage index. Stage 1 is where the raffinate exits and the barren organic enters.
Stage 25 is where the strip acid enters. **The organic flows from stage 1 toward 25;
the aqueous flows the opposite direction within each section.**

---

## Column 2 — Section Label (Column B)

**Formula (example for row 5, stage 1):**
```excel
=IF(A5<=INPUTS!$B$5, "Extraction",
   IF(A5<=INPUTS!$B$5+INPUTS!$B$6, "Scrub",
      "Strip"))
```

**What it does:**
Reads the stage number in column A and compares it to the stage counts in INPUTS.

- If stage ≤ n_ext → label is "Extraction"
- If stage ≤ n_ext + n_scr → label is "Scrub"
- Otherwise → label is "Strip"

**Why this matters:** Each section uses a different HCl normality, so the section
label controls which D value gets calculated in the next column.

---

## Column 3 — D_i : Distribution Coefficient (first column of each element block)

**Formula (example: Pr, stage 1, row 5):**
```excel
=IF(A5<=INPUTS!$B$5,
    INPUTS!$C$18 * INPUTS!$B$8 ^ INPUTS!$D$18,
    IF(A5<=INPUTS!$B$5+INPUTS!$B$6,
       INPUTS!$C$18 * INPUTS!$B$9 ^ INPUTS!$D$18,
       INPUTS!$C$18 * INPUTS!$B$10 ^ INPUTS!$D$18))
```

**What it calculates:**
```
D  =  P × [HCl]^Q
```

**Breaking it down piece by piece:**

| Part of formula | What it means |
|---|---|
| `INPUTS!$C$18` | P constant for Pr (from Distribution Coefficients table) |
| `INPUTS!$D$18` | Q exponent for Pr |
| `INPUTS!$B$8` | HCl normality in Extraction section |
| `INPUTS!$B$9` | HCl normality in Scrub section |
| `INPUTS!$B$10` | HCl normality in Strip section |

The IF selects which HCl normality to use based on which section this stage belongs to.

**Physical meaning:**
D tells you how much more of the element is in the organic phase versus the aqueous phase
at equilibrium. A high D means the organic strongly prefers that element. Because Q is
negative, increasing HCl always decreases D — high acid pushes elements back to aqueous.

---

## Column 4 — a : Sub-Diagonal Coefficient

**Formula:**
```excel
=IF(A5=1, 0, INPUTS!$B$13)
```

**What it calculates:**
```
a_1  =  0            (stage 1 has no stage before it)
a_n  =  O            (all other stages)
```

**Physical meaning:**
The `a` coefficient represents the organic flow **entering** from the previous stage.
Since the organic flows from stage 1 to 25, stage n receives organic from stage n−1.
Stage 1 receives barren (empty) organic — there is no stage 0 — so `a_1 = 0`.
For all other stages, the incoming organic flow is simply O (from INPUTS).

---

## Column 5 — b : Main Diagonal Coefficient

**Formula:**
```excel
=IF(A5 > INPUTS!$B$5 + INPUTS!$B$6,
    -INPUTS!$B$13 - INPUTS!$B$12 / C5,
    IF(A5 > INPUTS!$B$5,
       -INPUTS!$B$13 - INPUTS!$B$12 * INPUTS!$B$14 / C5,
       -INPUTS!$B$13 - (INPUTS!$B$12 * INPUTS!$B$14 + INPUTS!$B$11) / C5))
```

**What it calculates — three cases:**

| Section | Formula | Explanation |
|---|---|---|
| Strip (stage > n_ext+n_scr) | `−O − S/D` | Aqueous flow in strip = S |
| Scrub (stage > n_ext) | `−O − S×R/D` | Aqueous flow in scrub = S×R (reflux) |
| Extraction | `−O − (S×R + F)/D` | Aqueous flow in extraction = S×R + F (reflux + feed) |

**Physical meaning:**
The `b` coefficient represents everything **leaving** stage n — both organic and aqueous.
The organic leaving is O (goes to stage n+1). The aqueous leaving depends on which section:
in the strip section, the aqueous just carries flow S; in the scrub section, the scrub
solution carries S×R; in extraction, the total aqueous leaving is the scrub return plus
the feed that entered.

The division by D converts from aqueous to organic units (since x = y/D).

The result is always **negative** because it represents outflows (what leaves the stage).

---

## Column 6 — c : Super-Diagonal Coefficient

**Formula:**
```excel
=IF(A5 > INPUTS!$B$5 + INPUTS!$B$6,
    INPUTS!$B$12 / C6,
    IF(A5 >= INPUTS!$B$5,
       INPUTS!$B$12 * INPUTS!$B$14 / C6,
       (INPUTS!$B$12 * INPUTS!$B$14 + INPUTS!$B$11) / C6))
```

For the **last stage** (stage 25, row 29):
```excel
=0
```

**What it calculates — three cases:**

| Section | Formula | Explanation |
|---|---|---|
| Strip | `S / D_{n+1}` | Strip aqueous entering from next stage |
| Scrub + Feed cell | `S×R / D_{n+1}` | Scrub aqueous entering from next stage |
| Extraction (not feed) | `(S×R+F) / D_{n+1}` | Full aqueous from next stage |

Notice: `c` references `C6` (the D value of the **next** row), not `C5` (current row).
This is because the aqueous arrives from the next stage with that stage's D value.

**Physical meaning:**
The `c` coefficient represents the aqueous **entering** from the next stage (n+1).
The aqueous flows countercurrent to the organic — so stage n receives aqueous from stage n+1.
The last stage has no stage after it, so `c = 0`.

**Key detail:** The feed cell (stage n_ext) uses `S×R` not `S×R+F` for its c coefficient.
This is because the feed itself enters directly at stage n_ext and is captured in the
`d` column (RHS) — it does not propagate through the c coefficient to the next stage.

---

## Column 7 — d : Right-Hand Side (RHS)

**Formula:**
```excel
=IF(A5 = INPUTS!$B$5, -INPUTS!$C$26 * INPUTS!$B$11, 0)
```
*(example for Nd; each element uses its own feed concentration row)*

**What it calculates:**
```
d_n  =  0                            for all stages except the feed stage
d_B  =  −(Feed concentration) × F   at the feed stage (n = n_ext)
```

**Physical meaning:**
The right-hand side of the equation is zero everywhere except where an external stream
enters the system. The only external stream in this model is the aqueous feed, which
enters at stage n_ext (stage B in the PDF notation).

The **negative sign** comes from moving the feed term to the right side of the equation:
the feed concentration × feed flowrate is a known quantity, so it goes to the RHS.

If the feed concentration for an element is 0 (e.g., Pr has 0 g/L feed in the reference
case), then `d = 0` everywhere and that element will appear on organic due to... wait,
actually if feed=0, Pr should be 0 everywhere. In the reference case Pr feed is 0 so
it correctly produces 0 values throughout.

---

## Column 8 — c′ : Forward Sweep, Modified Super-Diagonal

This is the **first step of the TDMA solver**. The forward sweep eliminates the
sub-diagonal (a) so that the system becomes upper-triangular, making back-substitution
possible in the next step.

**Formula for Stage 1 (row 5) — base case:**
```excel
= c / b
= F5 / E5
```

**Formula for Stage 2 onward (row 6+):**
```excel
= c_n / (b_n − a_n × c′_{n−1})
= F6 / (E6 − D6 * H5)
```

**What it calculates:**
Define the intermediate denominator `w_n = b_n − a_n × c′_{n−1}`.
Then `c′_n = c_n / w_n`.

**Why this works:**
The Thomas Algorithm works by systematically eliminating the term `a_n × y_{n−1}` from
each equation. After processing row n, the sub-diagonal entry is gone and the equation
only involves `y_n` and `y_{n+1}`. The modified coefficient `c′` stores the new
super-diagonal after this elimination.

Each c′ depends only on the c′ from the row **above** — which is why you fill this
column **top to bottom**.

---

## Column 9 — d′ : Forward Sweep, Modified RHS

**Formula for Stage 1 (row 5) — base case:**
```excel
= d / b
= G5 / E5
```

**Formula for Stage 2 onward (row 6+):**
```excel
= (d_n − a_n × d′_{n−1}) / (b_n − a_n × c′_{n−1})
= (G6 − D6 * I5) / (E6 − D6 * H5)
```

**What it calculates:**
The same denominator `w_n` used in c′ is used here. The numerator subtracts the
contribution of the eliminated sub-diagonal term from the RHS.

**Physical meaning:**
After the forward sweep, the system has been converted from:
```
a·y_{n-1} + b·y_n + c·y_{n+1} = d
```
into an equivalent upper-triangular form:
```
y_n + c′·y_{n+1} = d′
```
The RHS `d′` carries the accumulated effect of all previous stages' feed terms.

**Important:** The feed term (`d = −Feed × F`) only appears at stage n_ext. After the
forward sweep, this feed information has been "distributed" upward through the d′ values
of all stages above the feed stage.

---

## Column 10 — y_org : Organic Concentration (the actual answer)

This is the **second step of TDMA** — back substitution. Starting from the last stage
and working upward.

**Formula for Stage 25 (row 29) — last stage:**
```excel
= d′
= I29
```

**Formula for all other stages (row 28 down to row 5):**
```excel
= d′_n − c′_n × y_{n+1}
= I28 − H28 * J29
```

**What it calculates:**
```
y_25  =  d′_25
y_n   =  d′_n  −  c′_n × y_{n+1}    (working from stage 24 down to stage 1)
```

**Physical meaning:**
`y_n` is the **organic phase concentration** of this element at stage n, in g/L.
This is what the model actually solves for. Each stage's organic concentration depends
on the stage below it (already computed), which is why you fill this column
**bottom to top**.

Excel handles this correctly without circular references because each cell references
the row below, which Excel resolves in the right order automatically.

**Reading the results:**
- `y_n` at stage 10 (the last extraction stage) tells you the loaded organic
- High `y` values in stages 1–10 = element is loading well onto organic
- Near-zero `y` values in stages 21–25 = element has been stripped off

---

## Aqueous Concentration Columns (x_Pr, x_Nd, x_Tb, x_Dy)

**Formula:**
```excel
= IF(D_i <> 0, y_org / D_i, 0)
```
*(example for Pr at stage 1: `=IF(C5<>0, J5/C5, 0)`)*

**What it calculates:**
```
x_n  =  y_n / D_n
```

**Physical meaning:**
From the equilibrium relationship `y = D × x`, we can rearrange to get `x = y / D`.
This gives the **aqueous phase concentration** of the element at each stage, in g/L.

This is what you would measure if you sampled the aqueous stream at stage n:
- `x` at stage 1 = raffinate concentration (what leaves with the raffinate)
- `x` at stage 21 = preg concentration (what exits as product from strip section)
- The IF guard prevents division-by-zero errors if D ever evaluates to 0.

---

## Total Organic Column (Σ y_org)

**Formula:**
```excel
= y_Pr + y_Nd + y_Tb + y_Dy
```
*(example row 5: `=J5 + R5 + Z5 + AH5`)*

**What it calculates:**
The total REE loading on the organic phase at each stage, summing all four elements.

**Key value to watch:** The total organic at stage n_ext (stage 10 = row 14) is the
"Loaded Organic" — this is what you see in the RESULTS sheet KEY METRIC. In the
reference case this should be **19.05 g/L**.

---

## Total Aqueous Column (Σ x_aq)

**Formula:**
```excel
= x_Pr + x_Nd + x_Tb + x_Dy
```

**What it calculates:**
The total REE concentration in the aqueous phase at each stage.

Useful for quickly seeing where the REE concentration peaks in the aqueous phase
(should peak in the strip section where elements are being stripped off the organic).

---

## How to Verify the Solution is Correct

After the model calculates, you can do a quick manual check on any stage using the
original mass balance equation:

```
a × y_{n-1}  +  b × y_n  +  c × y_{n+1}  should equal  d
```

For any interior stage (not feed stage), this should equal **zero**.
For the feed stage (stage n_ext), this should equal **−Feed × F**.

If any row gives a value significantly different from these, there is an error in
the coefficients (a, b, c) or the D values.

---

## Full Sequence Summary

```
INPUTS (P, Q, HCl, F, S, O, R, n_ext, n_scr, n_str, Feed concentrations)
    │
    ▼
D_i = P × [HCl_section]^Q          ← one value per stage per element
    │
    ▼
a, b, c, d coefficients             ← tridiagonal matrix row by row
    │
    ▼
c′, d′ forward sweep (top → bottom) ← eliminates sub-diagonal a
    │
    ▼
y_org back substitution (bottom → top) ← solves for organic concentrations
    │
    ▼
x_aq = y_org / D                   ← converts to aqueous concentrations
    │
    ▼
RESULTS: raffinate (stage 1), loaded organic (stage n_ext), preg (stage n_ext+n_scr+1)
```

---

*Solver Formulas Reference — Alind SX Steady-State Model, D2EHPA/HCl*
