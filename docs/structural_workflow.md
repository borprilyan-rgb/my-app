# Structural Workflow

## Purpose

The Structural detail page supports review of the Structural Rate used in Cost Analysis. It provides a structured place to enter structural detail assumptions under Area Analysis.

The detail page derives a suggested Structural Rate. It does not directly replace the final Structural total.

## Page Location

Area Analysis > Structural

## Current Calculation Principle

GBA is read from the existing GBA input and is non-editable inside the Structural page.

The final Structural project cost remains:

| Item | Formula |
| --- | --- |
| Structural Total | GBA x Structural Rate |

Cost Analysis owns the final project cost formula. The Structural detail page calculates a derived rate for review, but the final total continues to use GBA multiplied by the active Structural Rate.

## Structural Rate Updates

The active Structural Rate (`u_struc`) changes only through explicit user action:

| Action | Result |
| --- | --- |
| Manual input in Cost Analysis | Updates the active Structural Rate |
| Apply Structural Detail Rate button | Applies the derived detail rate to the active Structural Rate |

The Structural detail page does not silently overwrite `u_struc`.

## Detail Rate Reconciliation

The detail page reconciles the Structural detail amount back to a rate:

| Value | Calculation |
| --- | --- |
| Structural Detail Total | Sum(Qty x Rate) |
| Derived Structural Rate | Structural Detail Total / GBA |

This derived rate is used for review and can be applied only through the explicit Apply Structural Detail Rate button.

## Excel Export and Import

Structural Excel export/import uses one sheet:

| Sheet | Purpose |
| --- | --- |
| Structural | Structural detail-derived rate input |

Import behavior:

| Row Type | Import Rule |
| --- | --- |
| Detail rows | Imported as Structural detail input |
| Summary or total rows | Not imported as detail rows |

Summary and total rows are display/export rows. They must not become input detail rows during import.

## Structural Detail Formulas

| No | Item | Unit | Quantity Formula | Amount Formula |
| --- | --- | --- | --- | --- |
| 1 | Sub/Superstructure | m3 | GBA x structural_ratio | Qty x Rate |
| 2 | Bekisting | m2 | Sub/Superstructure x bekisting_ratio | Qty x Rate |
| 3 | Pembesian | kg | Sub/Superstructure x pembesian_ratio | Qty x Rate |
| 4 | Readymix Concrete | m3 | Sub/Superstructure x readymix_ratio | Qty x Rate |
| 5 | Rebar | kg | Sub/Superstructure x rebar_ratio x waste_factor | Qty x Rate |
| 6 | Prestress Works | ls | Manual input | Qty x Rate |
| 7 | Steelworks | kg | Manual input | Qty x Rate |
| 8 | Others | m3 | Sub/Superstructure | Qty x Rate |
