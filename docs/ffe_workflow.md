# FF&E Workflow

## Purpose

The FF&E detail page supports review of the FF&E Rate used in Cost Analysis. It provides a structured place to enter FF&E detail quantities and rates under Area Analysis.

The detail page derives a suggested FF&E Rate. It does not directly replace the final FF&E total.

## Page Location

Area Analysis > FF&E

## Current Calculation Principle

Rooms is read from the existing project input or calculation and is non-editable inside the FF&E page.

Rooms is a reminder and the denominator for the derived rate. Detail item quantities remain manual and editable.

The final FF&E project cost remains:

| Item | Formula |
| --- | --- |
| FF&E Total | Rooms x FF&E Rate |

Cost Analysis owns the final project cost formula. The FF&E detail page calculates a derived rate for review, but the final total continues to use Rooms multiplied by the active FF&E Rate.

## FF&E Rate Updates

The active FF&E Rate (`u_ffe`) changes only through explicit user action:

| Action | Result |
| --- | --- |
| Manual input in Cost Analysis | Updates the active FF&E Rate |
| Apply FF&E Detail Rate button | Applies the derived detail rate to the active FF&E Rate |

The FF&E detail page does not silently overwrite `u_ffe`.

## Detail Rate Reconciliation

The detail page reconciles the FF&E detail amount back to a rate:

| Value | Calculation |
| --- | --- |
| FF&E Detail Total | Sum(Qty x Rate) |
| Derived FF&E Rate | FF&E Detail Total / Rooms |

This derived rate is used for review and can be applied only through the explicit Apply FF&E Detail Rate button.

## Excel Export and Import

FF&E Excel export/import uses one sheet:

| Sheet | Purpose |
| --- | --- |
| FF&E | FF&E detail-derived rate input |

The FF&E sheet shows a Current Project Rooms reminder. This reminder is informational only and is not imported as a detail row.

Import behavior:

| Row Type | Import Rule |
| --- | --- |
| Detail rows | Imported as FF&E detail input |
| Summary, total, or reminder rows | Not imported as detail rows |

Summary, total, and reminder rows are display/export rows. They must not become input detail rows during import.

## FF&E Detail Items

| No. | Item | Unit | Quantity Formula | Editable Logic |
| --- | --- | --- | --- | --- |
| 1 | Seater & Chair | unit | Manual Qty | Qty E, Rate E |
| 2 | Beds & Linen | unit | Manual Qty | Qty E, Rate E |
| 3 | Kitchen Cabinet, Drawer | unit | Manual Qty | Qty E, Rate E |
| 4 | Electronic: TV 32", Minibar, Kettle, SDB | unit | Manual Qty | Qty E, Rate E |
| 5 | Housewares | unit | Manual Qty | Qty E, Rate E |
| 6 | Stove with 2 burner + Hoods | unit | Manual Qty | Qty E, Rate E |
| 7 | Microwave, Refrigerator, Washing Machine | unit | Manual Qty | Qty E, Rate E |
| 8 | Others: Artworks | unit | Manual Qty | Qty E, Rate E |
| 9 | Misc (Linen/Gym) | unit | Manual Qty | Qty E, Rate E |

## Input Notes

Rooms is not used to auto-fill detail quantities. Users may input quantities equal to Rooms when appropriate, but the app must not force it.

The derived rate uses Rooms only after the total detail amount has been calculated.
