<?php

use avadim\FastExcelReader\Excel;
use PHPUnit\Framework\TestCase;

class ColorsTest extends \PHPUnit\Framework\TestCase
{
    public function testDimension()
    {
        $file = __DIR__ . '/Files/colors.xlsx';
        self::assertFileExists($file);
        $excel = Excel::open($file);
        $sheet = $excel->sheet();

        $data = $sheet->readCells(true);

        $colors = [];
        foreach ($data as $cell) {
            $colors[] = $excel->getCompleteStyleByIdx($cell['s'])['fill']['fill-color'];
        }

        $checkColors = [
            '#F2F2F2',
            '#D9D9D9',
            '#BFBFBF',
            '#A6A6A6',
            '#808080',
            '#F3F2F2',
            '#D9D8D9',
            '#BFBFC0',
            '#A7A5A6',
            '#817F81',
        ];

        foreach ($colors as $key => $color) {
            $this->assertEquals($checkColors[$key], $color);
        }
    }
}