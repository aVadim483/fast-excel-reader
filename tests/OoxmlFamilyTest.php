<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Xlsx\XlsxBook;
use avadim\FastExcelReader\Tests\Support\TemporaryFilesTestCase;

final class OoxmlFamilyTest extends TemporaryFilesTestCase
{
    public static function formats(): array
    {
        return [['xlsx'], ['xlsm'], ['xltx'], ['xltm']];
    }

    /** @dataProvider formats */
    public function testRealExcelFilesThroughAllFactories(string $extension): void
    {
        $file = __DIR__ . '/test_files/ooxml/family.' . $extension;
        $reference = Excel::openXlsx(__DIR__ . '/test_files/ooxml/family.xlsx');
        $expected = $reference->sheet('ReaderData')->readCellsWithStyles();
        self::assertTrue(Excel::isXlsx($file));
        self::assertTrue(Excel::validate($file));
        self::assertFalse(Excel::isOds($file));
        $zip = new \ZipArchive();
        $zip->open($file);
        self::assertSame(in_array($extension, ['xlsm', 'xltm'], true), $zip->locateName('xl/vbaProject.bin') !== false);
        $zip->close();
        $stream = fopen($file, 'rb');
        foreach ([Excel::openXlsx($file), Excel::open($file), Excel::openString(file_get_contents($file)), Excel::openStream($stream)] as $book) {
            self::assertInstanceOf(XlsxBook::class, $book);
            self::assertSame($expected, $book->sheet('ReaderData')->readCellsWithStyles());
            self::assertTrue($book->sheet('ReaderData')->isMerged('B4'));
            self::assertSame('Second sheet', $book->sheet('Second')->readCells()['A1']);
            $book->close();
        }
        self::assertIsResource($stream);
        fclose($stream);
        $reference->close();
        $renamed = $this->temporaryFile(file_get_contents($file), 'input.csv');
        $book = Excel::open($renamed);
        self::assertInstanceOf(XlsxBook::class, $book);
        $book->close();
    }

    public function testXlsbIsRejectedWithSpecificMessage(): void
    {
        $file = $this->directory . '/binary.xlsx';
        $zip = new \ZipArchive();
        $zip->open($file, \ZipArchive::CREATE);
        $zip->addFromString('xl/workbook.bin', 'binary marker');
        $zip->close();
        self::assertFalse(Excel::isXlsx($file));
        $this->expectException(Exception::class);
        $this->expectExceptionMessage('XLSB');
        Excel::open($file);
    }

    public function testEmptyZipDoesNotFallThroughToCsv(): void
    {
        $file = $this->temporaryFile("PK\x05\x06" . str_repeat("\0", 18), 'empty.zip');
        self::assertFalse(Excel::isXlsx($file));
        $this->expectException(Exception::class);
        $this->expectExceptionMessage('Not an XLSX workbook');
        Excel::open($file);
    }
}
