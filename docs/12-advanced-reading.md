# Advanced Reading

[← Back to README](../README.md) | [Documentation index](../README.md#documentation) | [🇷🇺 Русский](ru/12-advanced-reading.md)

## Advanced example
```php
use \avadim\FastExcelReader\Excel;

$file = __DIR__ . '/files/demo-02-advanced.xlsx';

$excel = Excel::open($file);

$result = [
    'sheets' => $excel->getSheetNames() // get all sheet names
];

$result['#1'] = $excel
    // select sheet by name
    ->selectSheet('Demo1') 
    // select area with data where the first row contains column keys
    ->setReadArea('B4:D11', true)  
    // set date format
    ->setDateFormat('Y-m-d') 
    // set key for column 'C' to 'Birthday'
    ->readRows(['C' => 'Birthday']); 

// read other arrays with custom column keys
// and in this case we define range by columns only
$columnKeys = ['B' => 'year', 'C' => 'value1', 'D' => 'value2'];
$result['#2'] = $excel
    ->selectSheet('Demo2', 'B:D')
    ->readRows($columnKeys);

$result['#3'] = $excel
    ->setReadArea('F5:H13')
    ->readRows($columnKeys);
```
You can set read area by defined names in workbook. For example if workbook has defined name **Headers** with range **Demo1!$B$4:$D$4**
then you can read cells by this name

```php
$excel->setReadArea('Values');
$cells = $excel->readCells();
```
Note that since the value contains the sheet name, this sheet becomes the default sheet.

You can set read area in the sheet
```php
$sheet = $excel->getSheet('Demo1')->setReadArea('Headers');
$cells = $sheet->readCells();
```
But if you try to use this name on another sheet, you will get an error
```php
$sheet = $excel->getSheet('Demo2')->setReadArea('Headers');
// Exception: Wrong address or range "Values"

```

If necessary, you can fully control the reading process using the method ```readSheetCallback()``` with callback-function
```php
use \avadim\FastExcelReader\Excel;

$excel = Excel::open($file);

$result = [];
$excel->readCallback(function ($row, $col, $val) use(&$result) {
    // Any manipulation here
    $result[$row][$col] = (string)$val;

    // if the function returns true then data reading is interrupted  
    return false;
});
var_dump($result);
```

## See also

* [Reading Data](11-reading-data.md) — row by row, array keys, empty cells
* [Cell Value Types & Date Formatter](13-dates-and-types.md)
* [API Reference](90-api-reference.md)
