<?php

declare(strict_types=1);

namespace avadim\FastExcelReader;

use PHPUnit\Framework\TestCase;

final class FastExcelReaderTest extends TestCase
{
    public function testExcelReader()
    {
        $file = __DIR__ . '/../demo/files/demo-00-test.xlsx';

        $excel = Excel::open($file);

        $this->assertEquals('A1:C3', $excel->sheet()->dimension());

        $result = $excel->readCells();
        $this->assertTrue(isset($result['A1']) && $result['A1'] === 'number');
        $this->assertTrue(isset($result['B3']) && $result['B3'] === 'Word');

        $result = $excel->readRows();
        $this->assertEquals(count($result), $excel->sheet()->countRows());
        $this->assertTrue(isset($result['1']['A']) && $result['1']['A'] === 'number');
        $this->assertTrue(isset($result['3']['B']) && $result['3']['B'] === 'Word');

        $result = $excel->readColumns();
        $this->assertEquals(count($result), $excel->sheet()->countCols());
        $this->assertTrue(isset($result['A']['1']) && $result['A']['1'] === 'number');
        $this->assertTrue(isset($result['B']['3']) && $result['B']['3'] === 'Word');

        // Read rows and use the first row as column keys
        $result = $excel->readRows(true);
        $this->assertTrue(isset($result['2']['number']) && $result['2']['number'] === 111);
        $this->assertTrue(isset($result['3']['name']) && $result['3']['name'] === 'Word');

        $result = $excel->readRows(true, Excel::KEYS_SWAP);
        $this->assertTrue(isset($result['number']['2']) && $result['number']['2'] === 111);
        $this->assertTrue(isset($result['name']['3']) && $result['name']['3'] === 'Word');

        $result = $excel->readRows(false, Excel::KEYS_ZERO_BASED);
        $this->assertTrue(isset($result[0][0]) && $result[0][0] === 'number');
        $this->assertTrue(isset($result[2][1]) && $result[2][1] === 'Word');

        $result = $excel->readRows(false, Excel::KEYS_ONE_BASED);
        $this->assertTrue(isset($result[1][1]) && $result[1][1] === 'number');
        $this->assertTrue(isset($result[3][2]) && $result[3][2] === 'Word');

        $result = $excel->readRows(['A' => 'bee', 'B' => 'honey'], Excel::KEYS_FIRST_ROW | Excel::KEYS_ROW_ZERO_BASED);
        $this->assertTrue(isset($result[0]['bee']) && $result[0]['bee'] === 111);
        $this->assertTrue(isset($result[1]['honey']) && $result[1]['honey'] === 'Word');

        $file = __DIR__ . '/../demo/files/demo-02-advanced.xlsx';
        $excel = Excel::open($file);

        $result = $excel
            ->selectSheet('Demo2', 'B5:D13')
            ->readRows();
        $this->assertTrue(isset($result[5]['B']) && $result[5]['B'] === 2000);
        $this->assertTrue(isset($result[13]['D']) && round($result[13]['D']) === 104.0);

        $columnKeys = ['B' => 'year', 'C' => 'value1', 'D' => 'value2'];
        $result = $excel
            ->selectSheet('Demo2', 'B5:D13')
            ->readRows($columnKeys, Excel::KEYS_ONE_BASED);
        $this->assertTrue(isset($result[5]['year']) && $result[5]['year'] === 2004);
        $this->assertTrue(isset($result[9]['value1']) && $result[9]['value1'] === 674);

        $sheet = $excel->sheet('Demo2')->setReadArea('b4:d13');
        $result = $sheet->readCellsWithStyles();

        $this->assertEquals('Lorem', $result['C4']['v']);
        $this->assertEquals('thin', $result['C4']['s']['border']['border-left-style']);
    }
}

