# Reading Data

[← Back to README](../README.md) | [Documentation index](../README.md#documentation) | [🇷🇺 Русский](ru/11-reading-data.md)

* [Read values row by row in loop](#read-values-row-by-row-in-loop)
* [Keys in resulting arrays](#keys-in-resulting-arrays)
* [Empty cells & rows](#empty-cells--rows)

## Read values row by row in loop
```php
$sheet = $excel->sheet();
foreach ($sheet->nextRow() as $rowNum => $rowData) {
    // $rowData is array ['A' => ..., 'B' => ...]
    $addr = 'C' . $rowNum;
    if ($sheet->hasImage($addr)) {
        $sheet->saveImageTo($addr, $fullDirectoryPath);
    }
    // handling of $rowData here
    // ...
}

// OR
foreach ($sheet->nextRow() as $rowNum => $rowData) {
    // handling of $rowData here
    // ...
    // get image list from current row
    $imageList = $sheet->getImageListByRow();
    foreach ($imageList as $imageInfo) {
        $imageBlob = $sheet->getImageBlob($imageInfo['address']);
    }
}

// OR
foreach ($sheet->nextRow(['A' => 'One', 'B' => 'Two'], Excel::KEYS_FIRST_ROW) as $rowNum => $rowData) {
    // $rowData is array ['One' => ..., 'Two' => ...]
    // ...
}
```
NOTE: Every time you call the ```foreach ($sheet->nextRow() as $rowIndex => $row)``` loop, 
reading data starts from the first row.

But there is an alternative way to read row by row - using the readNextRow() method. 
In this case, you first need to call the ```$sheet->reset(...)``` method with the required reading parameters, 
and then you can call `````$sheet-readNextRow()`````. If at some point you need to start reading data from the beginning, 
you need to call ```$sheet->reset(...)``` again.

```php
// Init the internal read generator
$sheet->reset(['A' => 'One', 'B' => 'Two'], Excel::KEYS_FIRST_ROW);

// read the first row
$rowData = $sheet->readNextRow();
var_dump($rowData);

// Read the next 3 rows
for ($i = 0; $i < 3; $i++) {
    $rowData = $sheet->readNextRow();
    var_dump($rowData);
}

// Reset the internal generator and read all rows starting from the first one
$sheet->reset(['A' => 'One', 'B' => 'Two'], Excel::KEYS_FIRST_ROW);
$result = [];
while ($rowData = $sheet->readNextRow()) {
    $result[] = $rowData;
}
var_dump($result);
```

## Keys in resulting arrays
```php
// Read rows and use the first row as column keys
$result = $excel->readRows(true);

// The same, written declaratively
$result = $excel->sheet()->withHeader()->readRows();

// Skip the header row but name the columns yourself, in column order
$result = $excel->sheet()->withHeader(['col1', 'col2'])->readRows();
```
Names passed to `withHeader()` are positional: the first name goes to the first column of the read
area, so no column letters are involved and the same call works on a sheet whose data does not start
at `A1`. A shorter list renames only the columns it covers; the rest keep the name from the header row.

You will get this result:
```text
Array
(
    [2] => Array
        (
            ['col1'] => 111
            ['col2'] => 'aaa'
        )
    [3] => Array
        (
            ['col1'] => 222
            ['col2'] => 'bbb'
        )
)
```
The optional second argument specifies the result array keys
```php

// Rows and cols start from zero
$result = $excel->readRows(false, Excel::KEYS_ZERO_BASED);
```
You will get this result:
```text
Array
(
    [0] => Array
        (
            [0] => 'col1'
            [1] => 'col2'
        )
    [1] => Array
        (
            [0] => 111
            [1] => 'aaa'
        )
    [2] => Array
        (
            [0] => 222
            [1] => 'bbb'
        )
)
```
Allowed values of result mode

| mode options        | descriptions                                                                    |
|---------------------|---------------------------------------------------------------------------------|
| KEYS_ORIGINAL       | rows from '1', columns from 'A' (default)                                       |
| KEYS_ROW_ZERO_BASED | rows from 0                                                                     |
| KEYS_COL_ZERO_BASED | columns from 0                                                                  |
| KEYS_ZERO_BASED     | rows from 0, columns from 0 (same as KEYS_ROW_ZERO_BASED + KEYS_COL_ZERO_BASED) |
| KEYS_ROW_ONE_BASED  | rows from 1                                                                     |
| KEYS_COL_ONE_BASED  | columns from 1                                                                  |
| KEYS_ONE_BASED      | rows from 1, columns from 1 (same as KEYS_ROW_ONE_BASED + KEYS_COL_ONE_BASED)   |

Additional options that can be combined with result modes

| options         | descriptions                                 |
|-----------------|----------------------------------------------|
| KEYS_FIRST_ROW  | the same as _true_ in the first argument     |
| KEYS_RELATIVE   | index from top left cell of area (not sheet) |
| KEYS_SWAP       | swap rows and columns                        |

For example
```php

$result = $excel->readRows(['A' => 'bee', 'B' => 'honey'], Excel::KEYS_FIRST_ROW | Excel::KEYS_ROW_ZERO_BASED);
```
You will get this result:
```text
Array
(
    [0] => Array
        (
            [bee] => 111
            [honey] => 'aaa'
        )

    [1] => Array
        (
            [bee] => 222
            [honey] => 'bbb'
        )

)
```

## Empty cells & rows

The library already skips empty cells and empty rows by default. Empty cells are cells where nothing is written, 
and empty rows are rows where all cells are empty. If a cell contains an empty string, it is not considered empty. 
But you can change this behavior and skip cells with empty strings.

```php
$sheet = $excel->sheet();

// Skip empty cells and empty rows
foreach ($sheet->nextRow() as $rowNum => $rowData) {
    // handle $rowData
}

// Skip empty cells and cells with empty strings
foreach ($sheet->nextRow([], Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL) as $rowNum => $rowData) {
    // handle $rowData
}

// Skip empty cells and empty rows (rows containing only whitespace characters are also considered empty)
foreach ($sheet->nextRow([], Excel::TRIM_STRINGS | Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL) as $rowNum => $rowData) {
    // handle $rowData
}
```
Other way
```php
$sheet->reset([], Excel::TRIM_STRINGS | Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL);
$rowData = $sheet->readNextRow();
// do something

$rowData = $sheet->readNextRow();
// handle next row

// ...
```

## See also

* [Getting Started](10-getting-started.md)
* [Advanced Reading](12-advanced-reading.md) — read areas, defined names, callbacks
* [API Reference](90-api-reference.md)
