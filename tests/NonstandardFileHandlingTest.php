<?php

use avadim\FastExcelReader\Excel;
use PHPUnit\Framework\TestCase;

class NonstandardFileHandlingTest extends TestCase
{
    public function testNonStandardFile(): void
    {
        $standardFilePath = __DIR__ . '/test_files/standard-file.xlsx';
        self::assertFileExists($standardFilePath);
        $standardExcel = Excel::open($standardFilePath);
        $standardSheet = $standardExcel->sheet();
        self::assertNotNull($standardSheet);

        $nonStandardFilePath = __DIR__ . '/test_files/nonstandard-file.xlsx';
        self::assertFileExists($nonStandardFilePath);

        $nonStandardExcel = Excel::open($nonStandardFilePath);
        $nonStandardSheet = $nonStandardExcel->sheet();
        self::assertNotNull($nonStandardSheet);

        $specSymbolFilePath = __DIR__ . '/test_files/spec#name%sym _.xlsx';
        self::assertFileExists($specSymbolFilePath);
        $excel = Excel::open($specSymbolFilePath);
        $sheet = $excel->sheet();

        $cells = $sheet->readCells();
        self::assertEquals('qwerty string', $cells['B2']);
    }
}
