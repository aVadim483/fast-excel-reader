<?php

use avadim\FastExcelReader\Excel;
use PHPUnit\Framework\TestCase;

class DimensionTest extends \PHPUnit\Framework\TestCase
{
    public function testDimension()
    {
        // sheet properties contain incorrect information
        $file = __DIR__ . '/Files/wrong-dimension.xlsx';
        self::assertFileExists($file);
        $excel = Excel::open($file);
        $sheet = $excel->sheet();

        $this->assertEquals('D4', $sheet->dimension());
        $this->assertEquals('C3:E5', $sheet->actualDimension());

        $this->assertEquals(4, $sheet->minRow());
        $this->assertEquals(3, $sheet->minActualRow());

        $this->assertEquals(4, $sheet->maxRow());
        $this->assertEquals(5, $sheet->maxActualRow());

        $this->assertEquals(1, $sheet->countRows());
        $this->assertEquals(3, $sheet->countActualRows());

        $this->assertEquals('D', $sheet->minColumn());
        $this->assertEquals('C', $sheet->minActualColumn());

        $this->assertEquals('D', $sheet->maxColumn());
        $this->assertEquals('E', $sheet->maxActualColumn());

        $this->assertEquals(1, $sheet->countColumns());
        $this->assertEquals(3, $sheet->countActualColumns());

    }
}