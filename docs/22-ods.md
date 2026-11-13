# Reading ODS

[Русский](ru/22-ods.md) · [README](../README.md) · [5.0 migration](../UPGRADE.md)

```php
use avadim\FastExcelReader\Excel;

$book = Excel::openOds('data.ods'); // open(), openString(), openStream() also detect ODS
try {
    foreach ($book->sheet()->nextRow() as $number => $row) {
        // Process the row.
    }
} finally {
    $book->close();
}
```

ODS returns `Ods\OdsBook` and `Ods\OdsSheet`. The common read API supports sheet
selection, areas, headers, key modes, row/cell/column reads and row generators.
Values include strings, numbers, boolean, dates and saved formula results.
No formulas are calculated. Formula text (`f`) and style index (`s`) are null.

Dates become Unix timestamps, or use the configured date formatter. Timezone-free
dates use UTC; explicit offsets are honored. `dateFormatter(null)` or `false`
returns the original ISO date text because ODS has no Excel serial value.
`office:time-value` stays an ISO duration string such as `PT27H15M`, not a clock
timestamp or `DateInterval`. Original values are available as `o` with cell metadata.

Merged ranges and actual dimensions are supported. `dimension()` is initially an
empty sentinel, as for CSV; `actualDimension()`, `dimensionArray()`, count/min/max
methods and `getMergedCells()` scan the selected table and cache its metadata.
Empty repeated rows and cells are skipped without changing subsequent physical
coordinates. Nonempty repetitions are clipped to the read area before expansion.
The initial reader supports at most 1,048,576 rows and 16,384 columns; values outside
these bounds and invalid repetitions raise an exception.

All tables share `content.xml`: opening enumerates the sheets in one streaming pass,
and selecting a later table requires scanning past preceding content. Each iterator
has its own XML cursor. A row's selected cells occupy memory; full results from
`readRows()` and cached merge metadata can still be large.

The first reader does not expose ODS styles, images, hyperlinks, protection,
visibility settings or translated OpenFormula text. Shared style methods return
empty style data. Format-specific XLSX methods are not available on `OdsSheet`.
Encrypted ODS and embedded objects are not supported. DTD declarations are rejected.
An unknown ZIP package is never treated as CSV.
