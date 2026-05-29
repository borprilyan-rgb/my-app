# Earthworks Workflow

## Purpose

The Earthworks detail page supports review of the Earthworks cost rate used in Cost Analysis. It provides a structured place to enter and review Earthworks breakdown items under Area Analysis.

The detail page is used to derive or review the Earthwork Rate. It does not directly replace the final Earthworks total.

## Current Calculation Principle

The final Earthworks project cost remains:

| Item | Formula |
| --- | --- |
| Earthworks Total | GBA x Earthwork Rate |

Cost Analysis owns the final project cost formula. The Earthworks detail page can calculate a derived rate for review, but the final total continues to use GBA multiplied by the active Earthwork Rate.

## Earthwork Rate Updates

The active Earthwork Rate (`u_earth`) changes only through explicit user action:

| Action | Result |
| --- | --- |
| Manual input in Cost Analysis | Updates the active Earthwork Rate |
| Apply Earthworks Detail Rate button | Applies the derived detail rate to the active Earthwork Rate |

The Earthworks detail page does not silently overwrite `u_earth`.

## Detail Rate Reconciliation

The detail page reconciles the detailed Earthworks amount back to a rate:

| Value | Calculation |
| --- | --- |
| Derived Earthwork Rate | Detail Total / GBA |

This derived rate is used for review and can be applied only through the explicit Apply Earthworks Detail Rate button.

## Excel Export and Import

Earthworks Excel export/import supports the detail rows used by the Earthworks detail page.

Import behavior:

| Row Type | Import Rule |
| --- | --- |
| Detail rows | Imported as Earthworks detail input |
| Summary or total rows | Not imported as detail rows |

Summary and total rows are display/export rows. They must not become input detail rows during import.

## Compact Warning Location

The compact Earthworks warning is shown under:

Area Analysis > GBA Input > Excel Form

## Current Earthworks Breakdown

| Code | Item |
| --- | --- |
| 1.2.1 | Cut Fill |
| 1.2.2 | Dewatering |
| 1.2.3 | Soil Improvement |
| 1.2.4 | Shoring System |
| 1.2.5 | Others |
