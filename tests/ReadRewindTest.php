<?php

use avadim\FastExcelReader\Excel;
use PHPUnit\Framework\TestCase;

class ReadRewindTest extends \PHPUnit\Framework\TestCase
{
    public function testReadRewind()
    {
        $file = __DIR__ . '/Files/standard-file.xlsx';
        self::assertFileExists($file);
        $excel = Excel::open($file);
        $sheet = $excel->sheet();

        $data1 = [];
        foreach ($sheet->nextRow(false, Excel::KEYS_ORIGINAL) as $rowIndex => $row) {
            $data1[$rowIndex] = $row;
            if ($rowIndex >= 3) {
                break;
            }
        }

        $data2 = [];
        foreach ($sheet->nextRow(false, Excel::KEYS_ORIGINAL) as $rowIndex => $row) {
            $data2[$rowIndex] = $row;
            if ($rowIndex >= 3) {
                break;
            }
        }
        self::assertEquals($data1, $data2);

        $data3 = [];
        for($i = 1; $i <= 3; $i++) {
            $data3[$i] = $sheet->readNextRow();
        }
        self::assertEquals($data2, $data3);
        self::assertEquals('Invoice Date', $data3[1]['A']);
        $data3 = [];
        for($i = 1; $i <= 3; $i++) {
            $data3[$i] = $sheet->readNextRow();
        }
        self::assertNotEquals($data2, $data3);
        self::assertNotEquals('Invoice Date', $data3[1]['A']);

        $sheet->reset();
        $data3 = [];
        for($i = 1; $i <= 3; $i++) {
            $data3[$i] = $sheet->readNextRow();
        }
        self::assertEquals($data2, $data3);
        self::assertEquals('Invoice Date', $data3[1]['A']);
    }
}