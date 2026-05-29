# Architectural Workflow

## Purpose

The Architectural detail page supports review of the Architectural Rate used in Cost Analysis. It provides a structured place to enter architectural detail assumptions under Area Analysis.

The detail page derives a suggested Architectural Rate. It does not directly replace the final Architectural total.

## Page Location

Area Analysis > Architectural

## Current Calculation Principle

GFA is read from the existing GFA input or calculation and is non-editable inside the Architectural page.

The final Architectural project cost remains:

| Item | Formula |
| --- | --- |
| Architectural Total | GFA x Architectural Rate |

Cost Analysis owns the final project cost formula. The Architectural detail page calculates a derived rate for review, but the final total continues to use GFA multiplied by the active Architectural Rate.

## Architectural Rate Updates

The active Architectural Rate (`u_arch`) changes only through explicit user action:

| Action | Result |
| --- | --- |
| Manual input in Cost Analysis | Updates the active Architectural Rate |
| Apply Architectural Detail Rate button | Applies the derived detail rate to the active Architectural Rate |

The Architectural detail page does not silently overwrite `u_arch`.

## Detail Rate Reconciliation

The detail page reconciles the Architectural detail amount back to a rate:

| Value | Calculation |
| --- | --- |
| Architectural Detail Total | Sum(Qty x Rate) |
| Derived Architectural Rate | Architectural Detail Total / GFA |

This derived rate is used for review and can be applied only through the explicit Apply Architectural Detail Rate button.

## Excel Export and Import

Architectural Excel export/import uses one sheet:

| Sheet | Purpose |
| --- | --- |
| Architectural | Architectural detail-derived rate input |

Import behavior:

| Row Type | Import Rule |
| --- | --- |
| Detail rows | Imported as Architectural detail input |
| Summary or total rows | Not imported as detail rows |

Summary and total rows are display/export rows. They must not become input detail rows during import.

## Architectural Detail Formulas

| No. | Item | Unit | Quantity Formula | Editable Logic |
| --- | --- | --- | --- | --- |
| 1 | Basic Finishes Work | m2 | GFA | GFA N, Rate E |
| 2.1 | Aluminium Facade / Window Wall | m2 | Facade x Window Wall % | Facade N, % E, Rate E |
| 2.2 | Kisi2 Facade / Double Skin | m2 | Facade x Double Skin % | Facade N, % E, Rate E |
| 2.3 | Precast Facade | m2 | Facade x Precast % | Facade N, % E, Rate E |
| 3 | Pintu Kaca Dalam Ruangan | unit | Glass Door Qty | Qty N, Rate E |
| 4 | Railing Balkon | m' | Rooms x Railing Length per Room | Rooms N, length E, Rate E |
| 5 | Pintu Kayu | unit | Wooden Door Qty | Qty N, Rate E |
| 6 | Pintu Besi | unit | Steel Door Qty | Qty N, Rate E |
| 7 | Shower Screen | unit | Manual Qty | Qty E, Rate E |
| 8 | Marble / Door Jamb Lift | m' | Manual Qty | Qty E, Rate E |
| 9 | Interior - Main Lobby & Typical Lobby | m2 | Lobby Interior Area | Qty N, Rate E |
| 10 | Signage / Fixtures | ls | Manual Qty | Qty E, Rate E |
| 11 | Gondola | unit | Gondola Unit | Qty E, Rate E |
| 12 | Roof - Skylight | m2 | Skylight Area | Qty E, Rate E |
| 13.1 | Sanitary Fittings - T. Wanita | unit | Female Toilet Qty | Qty E, Rate E |
| 13.2 | Sanitary Fittings - T. Pria | unit | Male Toilet Qty | Qty E, Rate E |
| 13.3 | Sanitary Fittings - T. Disable | unit | Disabled Toilet Qty | Qty E, Rate E |
| 13.4 | Sanitary Fittings - Musholla | unit | Mushola Qty | Qty E, Rate E |
| 13.5 | Sanitary Fittings - Toilet Unit | unit | Rooms x Toilet Private per Room | Rooms N, toilet/room E, Rate E |
| 14 | Kitchen Equipment | unit | Rooms | Qty N, Rate E |
| 15.1 | Ironmongeries - Pintu Kayu | unit | Wooden Door Qty | Qty N, Rate E |
| 15.2 | Ironmongeries - Pintu Besi | unit | Steel Door Qty | Qty N, Rate E |
| 16.1 | Keramik & HT | m2 | GFA x HT % x Overlap x Waste | GFA N, % E, overlap E, waste E, Rate E |
| 16.2 | Marmer | m2 | GFA x Marmer % x Overlap x Waste | GFA N, % E, overlap E, waste E, Rate E |
| 16.3 | Vinyl | m2 | GFA x Vinyl % x Overlap x Waste | GFA N, % E, overlap E, waste E, Rate E |
| 17 | Carpet | m2 | Carpet Area | Qty E, Rate E |
| 18 | Kaca | m2 | Glass Area | Qty E, Rate E |

## Input Notes

Overlap and Waste default to 0. Example values such as 1.2 or 1.1 are examples only, not defaults.

Marble / Door Jamb Lift is manual quantity, not GFA x Marmer.
