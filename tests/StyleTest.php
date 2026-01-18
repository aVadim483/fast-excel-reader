<?php

use avadim\FastExcelReader\Excel;
use PHPUnit\Framework\TestCase;

class StyleTest extends \PHPUnit\Framework\TestCase
{
    public function testColors()
    {
        $file = __DIR__ . '/Files/styles.xlsx';
        self::assertFileExists($file);
        $excel = Excel::open($file);
        $sheet = $excel->sheet('colors');

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

    public function testColStyles()
    {
        $file = __DIR__ . '/Files/styles.xlsx';
        self::assertFileExists($file);
        $excel = Excel::open($file);
        $sheet = $excel->sheet('styles');

        $data = $sheet->getAllColAttributes();
        $this->assertEquals('15', $data['C']['style']);

        $data = $sheet->getColumnAttributes(3);
        $this->assertEquals('15', $data['style']);

        $data = $sheet->getColumnAttributes('3');
        $this->assertEquals('15', $data['style']);

        $data = $sheet->getColumnAttributes('c');
        $this->assertEquals('15', $data['style']);

        $data = $sheet->getColumnStyle(3, true);
        $this->assertEquals('#FFFF00', $data['fill-color']);

        $data = $sheet->getColumnStyle('C', true);
        $this->assertEquals('#FFFF00', $data['fill-color']);

        $data = $sheet->getAllRowAttributes();
        $this->assertEquals('3', $data['3']['r']);

        $data = $sheet->getRowStyle(4);
        $this->assertEquals('#202020', $data['font']['font-color']);
    }
}