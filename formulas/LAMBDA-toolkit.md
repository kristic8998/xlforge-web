# XLForge Formula Toolkit (no scripts needed)

Works in **any** Excel on the web — no Automate tab, no license requirements.
These are named LAMBDA functions: paste each definition once, then use it like a built-in function.

## How to install a function

1. In Excel on the web: **Formulas → Name Manager → New**
2. **Name**: the function name (e.g., `CLEANTEXT`)
3. **Refers to**: paste the LAMBDA definition
4. Click OK. Now `=CLEANTEXT(A2)` works anywhere in the workbook.

---

## CLEANTEXT — trim, de-space, remove junk characters

Fixes stray spaces, double spaces, non-breaking spaces (CHAR 160), and non-printing characters in one call.

```
=LAMBDA(txt, TRIM(CLEAN(SUBSTITUTE(txt, UNICHAR(160), " "))))
```

Usage: `=CLEANTEXT(A2)` → `"  gadget  pro "` becomes `"gadget pro"`

## PROPERCLEAN — clean + Proper Case in one step

```
=LAMBDA(txt, PROPER(TRIM(CLEAN(SUBSTITUTE(txt, UNICHAR(160), " ")))))
```

Usage: `=PROPERCLEAN(A2)` → `" bob SMITH "` becomes `"Bob Smith"`

## TEXTTODATE — convert text dates in mixed formats

Handles `2026-01-15`, `1/15/2026`, and `Jan 15 2026`. Returns a real date, or #VALUE! if unparseable.

```
=LAMBDA(txt,
  LET(t, TRIM(txt),
    IFERROR(DATEVALUE(t),
      IFERROR(DATEVALUE(SUBSTITUTE(t, " ", "/", 1)),
        DATEVALUE(TEXTBEFORE(t, " ") & " " & TEXTAFTER(t, " "))
      )
    )
  )
)
```

Usage: `=TEXTTODATE(B2)` then format the cell as Date (yyyy-mm-dd).

## CLEANRANGE — clean a whole column at once (dynamic array)

Returns a spilled, cleaned copy of an entire range — no fill-down needed.

```
=LAMBDA(rng, MAP(rng, LAMBDA(v, IF(ISTEXT(v), PROPER(TRIM(CLEAN(SUBSTITUTE(v, UNICHAR(160), " ")))), v))))
```

Usage: `=CLEANRANGE(A2:E201)` in an empty area → instantly spills a cleaned copy of the table.

## DEDUPED — remove exact duplicate rows

```
=LAMBDA(rng, UNIQUE(rng))
```

Usage: `=DEDUPED(A2:E201)` (yes, `UNIQUE` does the work — the named wrapper keeps your formulas self-documenting).

## GROUPSUM — instant summary table (group + total)

A one-formula pivot: unique groups in one column, their sums beside them.

```
=LAMBDA(groups, amounts,
  LET(u, SORT(UNIQUE(groups)),
    HSTACK(u, SUMIFS(amounts, groups, u))
  )
)
```

Usage: `=GROUPSUM(C2:C201, F2:F201)` → spills a Region | Total table. Select it and Insert → Chart for the visual.

---

## The full clean-and-report recipe

With the functions above installed, transforming a messy table is three formulas:

```
=CLEANRANGE(A2:F201)          ' in H2 — cleaned copy of the data
=DEDUPED(H2#)                 ' in P2 — duplicates removed (note the # spill reference)
=GROUPSUM(R2#, U2#)           ' summary table from the deduped columns you need
```
