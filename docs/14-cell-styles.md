# Cell Styles

[← Back to README](../README.md) | [Documentation index](../README.md#documentation) | [🇷🇺 Русский](ru/14-cell-styles.md)

## How to get complete info about the cell style

Usually read functions return just cell values, but you can read the values with styles.
In this case, for each cell, not a scalar value will be returned, but an array 
like ['v' => _scalar_value_, 's' => _style_array_, 'f' => _formula_]

```php
$excel = Excel::open($file);

$sheet = $excel->sheet();

$rows = $sheet->readRowsWithStyles();
$columns = $sheet->readColumnsWithStyles();
$cells = $sheet->readCellsWithStyles();

$cells = $sheet->readCellsWithStyles();
```
Or you can read styles only (without values)
```php
$cells = $sheet->readCellStyles();
/*
array (
  'format' => 
  array (
    'format-num-id' => 0,
    'format-pattern' => 'General',
  ),
  'font' => 
  array (
    'font-size' => '10',
    'font-name' => 'Arial',
    'font-family' => '2',
    'font-charset' => '1',
  ),
  'fill' => 
  array (
    'fill-pattern' => 'solid',
    'fill-color' => '#9FC63C',
  ),
  'border' => 
  array (
    'border-left-style' => NULL,
    'border-right-style' => NULL,
    'border-top-style' => NULL,
    'border-bottom-style' => NULL,
    'border-diagonal-style' => NULL,
  ),
)
 */
$cells = $sheet->readCellStyles(true);
/*
array (
  'format-num-id' => 0,
  'format-pattern' => 'General',
  'font-size' => '10',
  'font-name' => 'Arial',
  'font-family' => '2',
  'font-charset' => '1',
  'fill-pattern' => 'solid',
  'fill-color' => '#9FC63C',
  'border-left-style' => NULL,
  'border-right-style' => NULL,
  'border-top-style' => NULL,
  'border-bottom-style' => NULL,
  'border-diagonal-style' => NULL,
)
 */
```
But we do not recommend using these methods with large files

## See also

* [Cell Value Types & Date Formatter](13-dates-and-types.md)
* [Sheet Metadata](16-sheet-metadata.md)
* [API Reference](90-api-reference.md)
