<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Csv\CsvBook;
use avadim\FastExcelReader\Csv\CsvOptions;
use avadim\FastExcelReader\Csv\CsvReader;
use avadim\FastExcelReader\Tests\Support\TemporaryFilesTestCase;

final class FactoryOpeningTest extends TemporaryFilesTestCase
{
    /** @dataProvider xlsxFiles */
    public function testExplicitXlsxFactory(string $name): void
    {
        $file = __DIR__ . '/test_files/' . $name;
        $book = Excel::openXlsx($file);
        self::assertSame(Excel::class, get_class($book));
        self::assertSame(Excel::open($file)->readRows(), $book->readRows());
    }

    public static function xlsxFiles(): array
    {
        return [['standard-file.xlsx'], ['nonstandard-file.xlsx'], ['spec#name%sym _.xlsx']];
    }

    public function testXlsxWithoutExtension(): void
    {
        $file = $this->temporaryFile(file_get_contents(__DIR__ . '/test_files/standard-file.xlsx'), 'no-extension');
        self::assertSame(Excel::open($file)->readRows(), Excel::openXlsx($file)->readRows());
    }

    /** @dataProvider invalidXlsxFiles */
    public function testXlsxDoesNotDispatchToOtherFormats(string $kind): void
    {
        $file = $this->directory . '/input';
        if ($kind === 'csv') {
            file_put_contents($file, "a,b\n1,2\n");
        }
        elseif ($kind === 'xls') {
            copy(__DIR__ . '/test_files/xls/demo-00-test.xls', $file);
        }
        elseif ($kind !== 'missing') {
            $zip = new \ZipArchive();
            $zip->open($file, \ZipArchive::CREATE);
            $zip->addFromString($kind === 'docx' ? 'word/document.xml' : 'note.txt', 'text');
            $zip->close();
        }
        $this->expectException(Exception::class);
        Excel::openXlsx($file);
    }

    public static function invalidXlsxFiles(): array
    {
        return [['csv'], ['xls'], ['zip'], ['docx'], ['missing']];
    }

    /** @dataProvider csvOptions */
    public function testCsvFactoriesKeepOptionsAndTypes($options): void
    {
        $file = $this->temporaryFile("id;name\n1;Alice\n2;Bob\n", 'text.xlsx');
        $book = Excel::openCsvBook($file, $options);
        self::assertSame(CsvBook::class, get_class($book));
        self::assertSame(1, $book->countSheets());
        self::assertSame((new CsvBook($file, $options))->readRows(), $book->readRows());
        self::assertSame(Excel::open($file, ['format' => 'csv', 'delimiter' => ';'])->readRows(), $book->readRows());
        $reader = Excel::openCsvReader($file, $options);
        $legacy = Excel::openCsv($file, $options);
        self::assertSame(CsvReader::class, get_class($reader));
        self::assertSame($legacy->getCsvLine(), $reader->getCsvLine());
        self::assertSame($legacy->fromRow(2)->readRows(), $reader->fromRow(2)->readRows());
    }

    public static function csvOptions(): array
    {
        return [[null], [[]], [['delimiter' => ';']], [new CsvOptions(['delimiter' => ';'])]];
    }

    public function testEmptyCsvKeepsConstructorContract(): void
    {
        $file = $this->temporaryFile('');
        self::assertSame([], Excel::openCsvBook($file)->readRows());
        self::assertSame([], Excel::openCsvReader($file)->readRows());
        $this->expectException(Exception::class);
        Excel::open($file);
    }

    public function testCsvEncodingAndTolerantMode(): void
    {
        $file = $this->temporaryFile(mb_convert_encoding("id;name\n1;Привет\n", 'UTF-16LE', 'UTF-8'));
        foreach ([['delimiter' => ';', 'encoding' => 'UTF-16LE'], new CsvOptions(['delimiter' => ';', 'encoding' => 'UTF-16LE'])] as $options) {
            self::assertSame('Привет', Excel::openCsvBook($file, $options)->readRows()[2]['B']);
            self::assertSame(Excel::openCsv($file, $options)->readRows(), Excel::openCsvReader($file, $options)->readRows());
        }
        $file = $this->temporaryFile("id;name\n1;ab\"cd\n");
        $options = ['delimiter' => ';', 'mode' => CsvOptions::TOLERANT_MODE];
        self::assertSame('ab"cd', Excel::openCsvBook($file, $options)->readRows()[2]['B']);
        self::assertSame(Excel::openCsv($file, $options)->readRows(), Excel::openCsvReader($file, $options)->readRows());
        $this->expectException(\Throwable::class);
        Excel::openCsvBook($file, ['delimiter' => ';', 'mode' => CsvOptions::STRICT_MODE])->readRows();
    }
}
