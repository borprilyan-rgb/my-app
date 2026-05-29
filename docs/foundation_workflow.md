# Foundation Workflow

## Purpose

The Foundation detail page supports review of the Foundation Rate used in Cost Analysis. It provides a structured place to enter foundation detail quantities and rates under Area Analysis.

The detail page derives a suggested Foundation Rate. It does not directly replace the final Foundation total.

## Page Location

Area Analysis > Foundation

## Current Calculation Principle

GBA is read from the existing GBA input and is non-editable inside the Foundation page.

The final Foundation project cost remains:

| Item | Formula |
| --- | --- |
| Foundation Total | GBA x Foundation Rate |

Cost Analysis owns the final project cost formula. The Foundation detail page calculates a derived rate for review, but the final total continues to use GBA multiplied by the active Foundation Rate.

## Foundation Rate Updates

The active Foundation Rate (`u_found`) changes only through explicit user action:

| Action | Result |
| --- | --- |
| Manual input in Cost Analysis | Updates the active Foundation Rate |
| Apply Foundation Detail Rate button | Applies the derived detail rate to the active Foundation Rate |

The Foundation detail page does not silently overwrite `u_found`.

## Detail Rate Reconciliation

The detail page reconciles the Foundation detail amount back to a rate:

| Value | Calculation |
| --- | --- |
| Foundation Detail Total | Sum(Qty x Rate) |
| Derived Foundation Rate | Foundation Detail Total / GBA |

This derived rate is used for review and can be applied only through the explicit Apply Foundation Detail Rate button.

## Excel Export and Import

Foundation Excel export/import uses one sheet:

| Sheet | Purpose |
| --- | --- |
| Foundation | Foundation detail-derived rate input |

Import behavior:

| Row Type | Import Rule |
| --- | --- |
| Detail rows | Imported as Foundation detail input |
| Summary or total rows | Not imported as detail rows |

Summary and total rows are display/export rows. They must not become input detail rows during import.

## Foundation Detail Items

| No | Item | Unit | Quantity | Rate | Amount |
| --- | --- | --- | --- | --- | --- |
| 1 | Supply Tiang Pancang | m' | Manual input | Editable | Qty x Rate |
| 2 | Install Tiang Pancang | m' | Manual input | Editable | Qty x Rate |
