# Стили ячеек

[🇬🇧 English](../14-cell-styles.md) · [← К README](../../README.ru.md) · [Оглавление](../../README.ru.md#документация)

## Как получить полную информацию о стиле ячейки

Обычно функции чтения возвращают только значения ячеек, но можно читать значения вместе со стилями.
В этом случае для каждой ячейки будет возвращено не скалярное значение, а массив
вида ['v' => _скалярное_значение_, 's' => _массив_стиля_, 'f' => _формула_]

```php
$excel = Excel::open($file);

$sheet = $excel->sheet();

$rows = $sheet->readRowsWithStyles();
$columns = $sheet->readColumnsWithStyles();
$cells = $sheet->readCellsWithStyles();

$cells = $sheet->readCellsWithStyles();
```
Или можно читать только стили (без значений)
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
Но мы не рекомендуем использовать эти методы с большими файлами

## Смотрите также

* [Типы значений и форматирование дат](13-dates-and-types.md)
* [Метаданные листа](16-sheet-metadata.md)
* [Справочник API](../90-api-reference.md)
