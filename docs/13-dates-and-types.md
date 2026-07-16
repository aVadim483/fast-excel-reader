# Cell Value Types & Date Formatter

[← Back to README](../README.md) | [Documentation index](../README.md#documentation) | [🇷🇺 Русский](ru/13-dates-and-types.md)

* [Date Formatter](#date-formatter)
* [Cell value types](#cell-value-types)

## Date Formatter
By default, all datetime values returns as timestamp. But you can change this behavior using ```dateFormatter()```

![demo date](../demo/files/img2.jpg)
```php
$excel = Excel::open($file);
$sheet = $excel->sheet()->setReadArea('B5:D7');
$cells = $sheet->readCells();
echo $cells['C5']; // -2205187200

// If argument TRUE is passed, then all dates will be formatted as specified in cell styles
// IMPORTANT! The datetime format depends on the locale
$excel->dateFormatter(true);
$cells = $sheet->readCells();
echo $cells['C5']; // '14.02.1900'

// You can specify date format pattern
$excel->dateFormatter('Y-m-d');
$cells = $sheet->readCells();
echo $cells['C5']; // '1900-02-14'

// set date formatter function
$excel->dateFormatter(fn($value) => gmdate('m/d/Y', $value));
$cells = $sheet->readCells();
echo $cells['C5']; // '02/14/1900'

// returns DateTime instance
$excel->dateFormatter(fn($value) => (new \DateTime())->setTimestamp($value));
$cells = $sheet->readCells();
echo get_class($cells['C5']); // 'DateTime'

// custom manipulations with datetime values
$excel->dateFormatter(function($value, $format, $styleIdx) use($excel) {
    // get Excel format of the cell, e.g. '[$-F400]h:mm:ss\ AM/PM'
    $excelFormat = $excel->getFormatPattern($styleIdx);

    // get format converted for use in php functions date(), gmdate(), etc
    // for example the Excel pattern above would be converted to 'g:i:s A'
    $phpFormat = $excel->getDateFormatPattern($styleIdx);
    
    // and if you need you can get value of numFmtId for this cell
    $style = $excel->getCompleteStyleByIdx($styleIdx, true);
    $numFmtId = $style['format-num-id'];
    
    // do something and write to $result
    $result = gmdate($phpFormat, $value);

    return $result;
});
```
Sometimes, if a cell's format is specified as a date but does not contain a date, the library may misinterpret this value. To avoid this, you can disable date formatting

![demo date](../demo/files/img3.jpg)

Here, cell B1 contains the string "3.2" and cell B2 contains the date 2024-02-03, but both cells are set to the date format

```php
$excel = Excel::open($file);
// default mode
$cells = $sheet->readCells();
echo $cell['B1']; // -2208798720 - the library tries to interpret the number 3.2 as a timestamp
echo $cell['B2']; // 1706918400 - timestamp of 2024-02-03

// date formatter is on
$excel->dateFormatter(true);
$cells = $sheet->readCells();
echo $cell['B1']; // '03.01.1900'
echo $cell['B2']; // '3.2'

// date formatter is off
$excel->dateFormatter(false);
$cells = $sheet->readCells();
echo $cell['B1']; // '3.2'
echo $cell['B2']; // 1706918400 - timestamp of 2024-02-03

```

## Cell value types

The library tries to determine the types of cell values, and in most cases it does it right. 
Therefore, you get numeric or string values. Date values are returned as a timestamp by default.
But you can change this behavior by setting the date format (see the formatting options for the date() php function).

```php
$excel = Excel::open($file);
$result = $excel->readCells();
print_r($result);
```
The above example will output:
```text
Array
(
    [B2] => -2205187200
    [B3] => 6614697600
    [B4] => -6845212800
)
```
```php
$excel = Excel::open($file);
$excel->setDateFormat('Y-m-d');
$result = $excel->readCells();
print_r($result);
```
The above example will output:
```text
Array
(
    [B2] => '1900-02-14'
    [B3] => '2179-08-12'
    [B4] => '1753-01-31'
)
```

## See also

* [Cell Styles](14-cell-styles.md)
* [API Reference](90-api-reference.md)
