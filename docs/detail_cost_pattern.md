# Detail Cost Pattern

## Purpose

Earthworks is the first detail cost page pattern, Structural is the second implemented pattern, and Foundation is the third implemented pattern. The pattern separates detailed input and rate review from the final project cost formula.

## Ownership Model

| Area | Responsibility |
| --- | --- |
| Area Analysis | Owns detail input pages and detailed quantity/cost entry |
| Cost Analysis | Owns final project cost formulas and active cost rates |

This keeps detailed review work close to area inputs while keeping final cost logic in Cost Analysis.

## Pattern Rules

| Rule | Practical Meaning |
| --- | --- |
| Detail pages may calculate suggested or derived rates | A detail page can summarize detail rows into a rate for review |
| Applying a derived rate must always be explicit | Users must click an Apply button or take a clear equivalent action |
| Never silently overwrite cost rates | Detail calculations must not automatically change active Cost Analysis rates |
| Final formulas stay in Cost Analysis | Detail pages support the formula; they do not own it |

Earthworks, Structural, and Foundation all follow detail-derived-rate logic. Each detail page calculates a suggested rate from detail rows, then Cost Analysis decides whether that rate becomes active.

For Earthworks, the detail page calculates:

| Derived Value | Calculation |
| --- | --- |
| Earthwork Rate | Earthworks Detail Total / GBA |

For Structural, the detail page calculates:

| Derived Value | Calculation |
| --- | --- |
| Structural Rate | Structural Detail Total / GBA |

For Foundation, the detail page calculates:

| Derived Value | Calculation |
| --- | --- |
| Foundation Rate | Foundation Detail Total / GBA |

The final Earthworks, Structural, and Foundation costs remain calculated in Cost Analysis as:

| Final Cost | Calculation |
| --- | --- |
| Earthworks Total | GBA x Earthwork Rate |
| Structural Total | GBA x Structural Rate |
| Foundation Total | GBA x Foundation Rate |

## Excel Import and Export Pattern

Excel workflows should preserve input detail rows only.

| Row Type | Treatment |
| --- | --- |
| Detail rows | Exported and imported as editable detail inputs |
| Summary rows | Display/export rows only |
| Total rows | Display/export rows only |

Summary and total rows must not be imported as detail input rows.

## Future Candidates

Future detail cost pages can follow the Earthworks, Structural, and Foundation pattern:

| Candidate |
| --- |
| Shoring |
| External works |
| Residential detail |
| Door/detail items |
