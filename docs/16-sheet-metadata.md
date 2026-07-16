# Sheet Metadata

[← Back to README](../README.md) | [Documentation index](../README.md#documentation) | [🇷🇺 Русский](ru/16-sheet-metadata.md)

* [Retrieve data validation rules](#retrieve-data-validation-rules)
* [Column Widths](#column-widths)
* [Row Heights](#row-heights)
* [Freeze Pane Info](#freeze-pane-info)
* [Tab Color Info](#tab-color-info)
* [Info about merged cells](#info-about-merged-cells)
* [Count rows and columns](#count-rows-and-columns)

## Retrieve data validation rules
Every sheet in your XLSX file can contain a set of data validation rules. To retrieve them, you can imply call `getDataValidations` on your sheet

```php
$excel = Excel::open($file);

$sheet = $excel->sheet();

$validations = $sheet->getDataValidations();
/*
[
  [
    'type' => 'list',
    'sqref' => 'E2:E527',
    'formula1' => '"Berlin,Cape Town,Mexico City,Moscow,Sydney,Tokyo"',
    'formula2' => null, 
  ], [
    'type' => 'decimal',
    'sqref' => 'G2:G527',
    'formula1' => '0.0',
    'formula2' => '999999.0',
  ],
]
*/
```

## Column Widths
Retrieve the width of a specific column in a sheet:

```php
$excel = Excel::open($file);
$sheet = $excel->selectSheet('SheetName');

// Get the width of column 1 (column 'A')
$columnWidth = $sheet->getColumnWidth(1);

echo $columnWidth; // Example: 11.85
```

## Row Heights
Retrieve the height of a specific row in a sheet:

```php
$excel = Excel::open($file);
$sheet = $excel->selectSheet('SheetName');

// Get the height of row 1
$rowHeight = $sheet->getRowHeight(1);

echo $rowHeight; // Example: 15
```

## Freeze Pane Info
Retrieve the freeze pane info for a sheet:

```php
$excel = Excel::open($file);
$sheet = $excel->selectSheet('SheetName');

// Get the freeze pane configuration
$freezePaneConfig = $sheet->getFreezePaneInfo();

print_r($freezePaneConfig);
/*
Example Output:
Array
(
    [xSplit] => 0
    [ySplit] => 1
    [topLeftCell] => 'A2'
)
*/
```

## Tab Color Info
Retrieve the tab color info for a sheet:

```php
$excel = Excel::open($file);
$sheet = $excel->selectSheet('SheetName');

// Get the tab color configuration
$tabColorConfig = $sheet->getTabColorInfo();

print_r($tabColorConfig);
/*
Example Output:
Array
(
    [theme] => '2'
    [tint] => '-0.499984740745262'
)
*/
```

## Info about merged cells

You can use the following methods:

* ```Sheet::getMergedCells()``` -- Returns all merged ranges
* ```Sheet::isMerged(string $cellAddress)``` -- Checks if a cell is merged
* ```Sheet::mergedRange(string $cellAddress)``` -- Returns merge range of specified cell

For example
```php
if ($sheet->isMerged('B3')) {
    $range = $sheet->mergedRange('B3');
}
```

## Count rows and columns

Each sheet contains the ```dimension``` property with the range of the area in which the data is written. 
If only one cell is filled on the sheet, then there should be an address of only this cell of the form "B2", 
otherwise it is a range of the form "B2:E10".

There are several methods that get data from this property:
* ```dimension()``` -- Returns dimension of default work area from sheet properties
* ```countRows()``` -- Count rows from dimension
* ```countColumns()``` -- Count columns from dimension
* ```minRow()``` -- The minimal row number from sheet properties
* ```maxRows()``` -- The maximal row number from sheet properties
* ```minColumn()``` -- The minimal column letter from sheet properties
* ```maxColumn()``` -- The maximal column letter from sheet properties

But sometimes the ```dimension``` property contains incorrect information. 
For example, it may contain the address of only the first cell of the data range or the address of only the last cell. 
In such cases, you can use methods that scan the entire sheet and count the actual number of rows and columns with data on the sheet.

IMPORTANT: these methods are slower than methods using the ```dimension``` property

* ```actualDimension()``` -- Returns dimension of the actual work area
* ```countActualRows()``` -- Count actual rows from the sheet
* ```minActualRow()``` -- The minimal actual row number
* ```maxActualRow()``` -- The maximal actual row number
* ```countActualColumns()``` -- Count actual columns from the sheet
* ```minActualColumn()``` -- The minimal actual column letter
* ```maxActualColumn()``` -- The maximal actual column letter

## See also

* [Cell Styles](14-cell-styles.md)
* [Full API Reference — Class Sheet](92-api-class-sheet.md)
