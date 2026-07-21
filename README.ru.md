[![GitHub Release](https://img.shields.io/github/v/release/aVadim483/fast-excel-reader)](https://packagist.org/packages/avadim/fast-excel-reader)
[![Packagist Downloads](https://img.shields.io/packagist/dt/avadim/fast-excel-reader?color=%23aa00aa)](https://packagist.org/packages/avadim/fast-excel-reader)
[![GitHub License](https://img.shields.io/github/license/aVadim483/fast-excel-reader)](https://packagist.org/packages/avadim/fast-excel-reader)
[![Static Badge](https://img.shields.io/badge/php-%3E%3D7.4-005fc7)](https://packagist.org/packages/avadim/fast-excel-reader)

# FastExcelReader

[🇬🇧 English](README.md) · 🇷🇺 Русский

**FastExcelReader** — часть проекта FastExcelPhp, в который входят

* [FastExcelWriter](https://packagist.org/packages/avadim/fast-excel-writer) — создание Excel-таблиц
* [FastExcelReader](https://packagist.org/packages/avadim/fast-excel-reader) — чтение Excel-таблиц и CSV-файлов
* [FastExcelTemplator](https://packagist.org/packages/avadim/fast-excel-templator) — генерация Excel-таблиц из XLSX-шаблонов
* [FastExcelLaravel](https://packagist.org/packages/avadim/fast-excel-laravel) — специальная редакция для **Laravel**

## Введение

Библиотека спроектирована лёгкой, очень быстрой и требует минимум памяти.

**FastExcelReader** умеет читать Excel-совместимые таблицы в формате XLSX (Office 2007+),
в устаревшем формате XLS (Office 97-2003) и CSV-файлы. Она только читает данные, но делает
это очень быстро и с минимальным расходом памяти.

## Возможности

### Поддержка формата XLSX

* Поддержка формата XLSX (Office 2007+) с несколькими листами
* Автоопределение типов валюта/число/дата
* Автоматическое и пользовательское форматирование значений даты-времени
* Библиотека умеет находить и извлекать изображения из XLSX-файлов
* Библиотека умеет читать параметры оформления ячеек — шаблоны форматирования, цвета, границы, шрифты и т.д.

### Поддержка формата XLS

* Поддержка устаревшего формата XLS (Office 97-2003, BIFF8) с тем же API, что и у XLSX
* Значения, даты, стили ячеек, текст формул и изображения возвращаются в той же форме
* Формат определяется по сигнатуре файла, а не по расширению
* Тоже потоковое чтение: лист читается за один проход вперёд с постоянным расходом памяти

### Поддержка формата CSV

* Определение разделителя: автоматическое или указанное вручную
* Широкая поддержка кодировок: любая кодировка, поддерживаемая вашим окружением PHP
* Поддержка полей с кавычками и без, а также экранированных кавычек (удвоенных)
* Многострочные поля: обработка переносов строк внутри полей в кавычках
* Два режима — строгий (следует RFC 4180) и толерантный (для нестандартных CSV-файлов)

## Установка

Используйте `composer`, чтобы установить **FastExcelReader** в свой проект:

```
composer require avadim/fast-excel-reader
```

## Быстрый старт

```php
use \avadim\FastExcelReader\Excel;

// И XLSX, и устаревший XLS открываются так же — ридер выбирается по сигнатуре файла
$excel = Excel::open(__DIR__ . '/files/demo.xlsx');

// Прочитать все строки в двумерный массив (СТРОКА x КОЛОНКА)
$rows = $excel->readRows();

// Или читать строку за строкой (экономно по памяти)
$sheet = $excel->sheet();
foreach ($sheet->nextRow() as $rowNum => $rowData) {
    // обработка $rowData
}
```

Подробнее — в руководстве [Начало работы](docs/ru/10-getting-started.md).

## Документация

### Руководства

* [Начало работы](docs/ru/10-getting-started.md) — чтение ячеек, строк и колонок
* [Чтение данных](docs/ru/11-reading-data.md) — построчно, ключи массивов, пустые ячейки и строки
* [Продвинутое чтение](docs/ru/12-advanced-reading.md) — области чтения, именованные диапазоны, колбэки
* [Типы значений и форматирование дат](docs/ru/13-dates-and-types.md)
* [Стили ячеек](docs/ru/14-cell-styles.md) — чтение полной информации о стиле ячеек
* [Изображения](docs/ru/15-images.md) — извлечение изображений из XLSX
* [Метаданные листа](docs/ru/16-sheet-metadata.md) — проверка данных, ширины колонок, высоты строк, закрепление областей, цвет вкладки, объединённые ячейки, размерность
* [Разбор CSV](docs/ru/20-csv.md) — чтение CSV-файлов
* [XLS (Excel 97-2003)](docs/ru/21-xls.md) — чтение устаревших книг XLS

### Справочник API

Справочник API генерируется автоматически из PHPDoc и доступен на английском языке:

* [API Reference](docs/90-api-reference.md)
  * [Class Excel](docs/91-api-class-excel.md)
  * [Class Sheet](docs/92-api-class-sheet.md)
  * [Class Csv\CsvReader](docs/94-api-class-csv-reader.md)
  * [Class Csv\CsvOptions](docs/95-api-class-csv-options.md)

### Примеры

Больше примеров — в папке [*/demo*](demo).

## Хотите поддержать FastExcelReader?

Если этот пакет оказался полезным, вы можете поставить мне звезду на GitHub.

Или можете сделать донат :)
* USDT (TRC20) TSsUFvJehQBJCKeYgNNR1cpswY6JZnbZK7
* USDT (ERC20) 0x5244519D65035aF868a010C2f68a086F473FC82b
* ETH 0x5244519D65035aF868a010C2f68a086F473FC82b
