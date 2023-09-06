<?php

declare(strict_types=1);

namespace avadim\FastExcelReader;

use PHPUnit\Framework\TestCase;

final class FastExcelReaderTest extends TestCase
{
    const DEMO_DIR = __DIR__ . '/../demo/files/';

    public function testExcelReader00()
    {
        // =====================
        $file = self::DEMO_DIR . 'demo-00-test.xlsx';
        $excel = Excel::open($file);

        $this->assertEquals('A1:C4', $excel->sheet()->dimension());

        $result = $excel->readCells();
        $this->assertTrue(isset($result['A1']) && $result['A1'] === 'name');
        $this->assertTrue(isset($result['B3']) && $result['B3'] === 6614697600);

        $result = $excel->readRows();
        $this->assertEquals(count($result), $excel->sheet()->countRows());
        $this->assertTrue(isset($result['1']['A']) && $result['1']['A'] === 'name');
        $this->assertTrue(isset($result['3']['B']) && $result['3']['B'] === 6614697600);

        $result = $excel->readColumns();
        $this->assertEquals(count($result), $excel->sheet()->countCols());
        $this->assertTrue(isset($result['A']['1']) && $result['A']['1'] === 'name');
        $this->assertTrue(isset($result['B']['4']) && $result['B']['4'] === -6845212800);

        // Read rows and use the first row as column keys
        $result = $excel->readRows(true);
        $this->assertTrue(isset($result['2']['name']) && $result['2']['name'] === 'James Bond');
        $this->assertTrue(isset($result['3']['birthday']) && $result['3']['birthday'] === 6614697600);

        $result = $excel->readRows(true, Excel::KEYS_SWAP);
        $this->assertTrue(isset($result['name']['3']) && $result['name']['3'] === 'Ellen Louise Ripley');

        $result = $excel->readRows(false, Excel::KEYS_ZERO_BASED);
        $this->assertTrue(isset($result[3][0]) && $result[3][0] === 'Captain Jack Sparrow');

        $result = $excel->readRows(false, Excel::KEYS_ONE_BASED);
        $this->assertTrue(isset($result[1][1]) && $result[1][1] === 'name');

        $result = $excel->readRows(['A' => 'Hero', 'C' => 'Secret']);
        $this->assertFalse(isset($result[0]['Hero']));
        $this->assertTrue(isset($result[1]['Hero']) && $result[1]['Hero'] === 'name');
        $this->assertTrue(isset($result[2]['Hero']) && $result[2]['Hero'] === 'James Bond');
        $this->assertTrue(isset($result[2]['B']) && $result[2]['B'] === -2205187200);
        $this->assertTrue(isset($result[2]['Secret']) && $result[2]['Secret'] === 4573);

        $result = $excel->readRows(['A' => 'Hero', 'C' => 'Secret'], Excel::KEYS_FIRST_ROW);
        $this->assertFalse(isset($result[0]['Hero']));
        $this->assertFalse(isset($result[1]['Hero']));
        $this->assertTrue(isset($result[2]['Hero']) && $result[2]['Hero'] === 'James Bond');
        $this->assertTrue(isset($result[2]['birthday']) && $result[2]['birthday'] === -2205187200);

        $result = $excel->readRows(['A' => 'Hero', 'C' => 'Secret'], Excel::KEYS_FIRST_ROW | Excel::KEYS_ROW_ZERO_BASED);
        $this->assertTrue(isset($result[0]['Hero']) && $result[0]['Hero'] === 'James Bond');
        $this->assertTrue(isset($result[0]['birthday']) && $result[0]['birthday'] === -2205187200);

        $result = $excel->readRows(['A' => 'Hero', 'C' => 'Secret'], Excel::KEYS_ROW_ZERO_BASED);
        $this->assertTrue(isset($result[0]['Hero']) && $result[0]['Hero'] === 'name');
        $this->assertTrue(isset($result[0]['B']) && $result[0]['B'] === 'birthday');
        $this->assertTrue(isset($result[1]['Hero']) && $result[1]['Hero'] === 'James Bond');
        $this->assertTrue(isset($result[1]['B']) && $result[1]['B'] === -2205187200);

        $result = $excel->readRows([]);
        $this->assertTrue(isset($result[1]['A']) && $result[1]['A'] === 'name');

        $result = $excel->readRows([], Excel::KEYS_FIRST_ROW);
        $this->assertTrue(isset($result[2]['name']) && $result[2]['name'] === 'James Bond');

        $result = [];
        $sheet = $excel->setReadArea('b2');
        foreach ($sheet->nextRow() as $row => $rowData) {
            $result[$row] = $rowData;
        }
        $this->assertCount(3, $result);
        $this->assertFalse(isset($result[1]));
        $this->assertFalse(isset($result[2]['A']));
        $this->assertTrue(isset($result[2]['B']) && $result[2]['B'] === -2205187200);
        $this->assertTrue(isset($result[2]['C']) && $result[2]['C'] === 4573);
        $this->assertFalse(isset($result[2]['D']));
        $this->assertFalse(isset($result[5]));

        $excel->setDateFormat('Y-m-d');
        $result = $excel->readCells();
        $this->assertEquals('1900-02-14', $result['B2']);
        $this->assertEquals('2179-08-12', $result['B3']);
        $this->assertEquals('1753-01-31', $result['B4']);
    }

    public function testExcelReader01()
    {
        // =====================
        $file = self::DEMO_DIR . 'demo-01-base.xlsx';
        $excel = Excel::open($file);

        $cells = $excel->sheet()->readCells(true);
        $this->assertEquals('A1', array_key_first($cells));
        $this->assertCount(4216, $cells);
        $this->assertEquals(142408, $cells['H2']['v']);

        $cells = $excel->sheet()->setReadArea('c10')->readCells();
        $this->assertEquals('C10', array_key_first($cells));
        $this->assertCount(3108, $cells);

        $cells = $excel->selectSheet('report', 'd10:e18')->readCells();
        $this->assertEquals('D10', array_key_first($cells));
        $this->assertCount(18, $cells);

    }

    public function testExcelReader02()
    {
        $file = self::DEMO_DIR . 'demo-02-advanced.xlsx';
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

        // default sheet is Demo2
        $sheet = $excel->getSheet()->setReadArea('b4:d13');
        $result = $sheet->readCellsWithStyles();

        $this->assertEquals('Lorem', $result['C4']['v']);
        $this->assertEquals('thin', $result['C4']['s']['border']['border-left-style']);

        $excel->selectSheet('Demo1');
        $this->assertEquals('Demo1', $excel->sheet()->name());

        $excel->selectSheet('Demo2');
        $this->assertEquals('Demo2', $excel->sheet()->name());

        $sheet = $excel->sheet('WrongSheet');
        $this->assertEquals(null, $sheet);

        $sheet = $excel->getSheet('Demo2', 'B4:D13', true);
        $result = $sheet->readRows();
        $this->assertTrue(isset($result[5]['Year']) && $result[5]['Year'] === 2000);
        $this->assertTrue(isset($result[5]['Lorem']) && $result[5]['Lorem'] === 235);

        $sheet = $excel->getSheet('Demo2', 'b:c');
        $result = $sheet->readRows();
        $this->assertTrue(isset($result[6]['B']) && $result[6]['B'] === 2001);
        $this->assertFalse(isset($result[6]['D']));

        $sheet = $excel->getFirstSheet();
        $result = $sheet->readRows(false, Excel::KEYS_ZERO_BASED);
        $this->assertTrue(isset($result[3][0]) && $result[3][0] === 'Giovanni');

        $this->assertEquals('Demo2', $excel->sheet()->name());

        $excel->setReadArea('Values');
        $result = $excel->readCells();
        $this->assertEquals('Giovanni', $result['B5']);

        $sheet = $excel->getSheet('Demo1')->setReadArea('Headers');
        $result = $sheet->readCells();
        $this->assertEquals('Name', $result['B4']);

        $this->expectException(\avadim\FastExcelReader\Exception::class);
        $excel->getSheet('Demo2')->setReadArea('Values');
    }

    public function testExcelReader03()
    {
        $file = self::DEMO_DIR . 'demo-03-images.xlsx';
        $excel = Excel::open($file);
        $this->assertEquals(2, $excel->countImages());

        $this->assertFalse($excel->sheet()->hasImage('c1'));
        $this->assertTrue($excel->sheet()->hasImage('c2'));

        $result = $excel->getImageList();
        $this->assertTrue(isset($result['Sheet1']['C2']));
        $this->assertEquals('image1.jpeg', $result['Sheet1']['C2']['file_name']);
    }

    public function testExcelReader04()
    {
        $file = self::DEMO_DIR . 'demo-04-styles.xlsx';
        $excel = Excel::open($file);
        $cells = $excel->readCellsWithStyles();
        $this->assertEquals('#9FC63C', $cells['A1']['s']['fill']['fill-color']);
        $this->assertEquals([
            'border-left-style' => 'thick',
            'border-right-style' => 'thin',
            'border-top-style' => 'thick',
            'border-bottom-style' => 'thin',
            'border-diagonal-style' => null,
            'border-left-color' => '#000000',
            'border-right-color' => '#000000',
            'border-top-color' => '#000000',
            'border-bottom-color' => '#000000',
        ], $cells['A6']['s']['border']);

        $cells = $excel->readCellStyles(true);
        $this->assertEquals([
            'format-num-id' => 0,
            'format-pattern' => 'General',
            'format-category' => 'general',
            'font-size' => '10',
            'font-name' => 'Arial',
            'font-family' => '2',
            'font-charset' => '1',
            'fill-pattern' => 'solid',
            'fill-color' => '#9FC63C',
            'border-left-style' => null,
            'border-right-style' => null,
            'border-top-style' => null,
            'border-bottom-style' => null,
            'border-diagonal-style' => null,
        ], $cells['A1']);
        $this->assertEquals('thick', $cells['A6']['border-left-style']);
        $this->assertEquals('thin', $cells['A6']['border-bottom-style']);
        $this->assertEquals('#000000', $cells['A6']['border-top-color']);
    }

    public function testDateFormatter()
    {
        // =====================
        $file = self::DEMO_DIR . 'demo-02-advanced.xlsx';
        $excel = Excel::open($file);

        $cells = $excel->sheet()->readCells();
        $this->assertEquals(18316800, $cells['C5']);
        $this->assertEquals(-777600, $cells['C6']);
        $this->assertEquals(-62121600, $cells['C7']);
        $this->assertEquals(38707200, $cells['C8']);

        $excel->setDateFormat('Y-m-d');
        $cells = $excel->sheet()->readCells();
        $this->assertEquals('1970-08-01', $cells['C5']);
        $this->assertEquals('1969-12-23', $cells['C6']);
        $this->assertEquals('1968-01-13', $cells['C7']);
        $this->assertEquals('1971-03-25', $cells['C8']);

        $excel->dateFormatter(fn($value) => gmdate('m/d/Y', $value));
        $cells = $excel->sheet()->readCells();
        $this->assertEquals('08/01/1970', $cells['C5']);
        $this->assertEquals('12/23/1969', $cells['C6']);
        $this->assertEquals('01/13/1968', $cells['C7']);
        $this->assertEquals('03/25/1971', $cells['C8']);

        $excel->dateFormatter(fn($value) => (new \DateTime())->setTimestamp($value)->format('z'));
        $cells = $excel->sheet()->readCells();
        $this->assertEquals('212', $cells['C5']);
        $this->assertEquals('356', $cells['C6']);
        $this->assertEquals('12', $cells['C7']);
        $this->assertEquals('83', $cells['C8']);

        $file = self::DEMO_DIR . 'demo-05-datetime.xlsx';
        $excel = Excel::open($file);

        $sheet = $excel->sheet()->setReadArea('B2:B2');
        $cells = $sheet->readCells();
        $this->assertEquals(441696063, $cells['B2']);

        $excel->dateFormatter(true);
        $cells = $sheet->readCells();
        $this->assertEquals('1983-12-31 05:21:03', $cells['B2']);

        $excel->dateFormatter('Y-m-d');
        $cells = $sheet->readCells();
        $this->assertEquals('1983-12-31', $cells['B2']);

        $excel->dateFormatter(fn($v) => gmdate('d/m/y', $v));
        $cells = $sheet->readCells();
        $this->assertEquals('31/12/83', $cells['B2']);
    }

    public function testFillRow()
    {
        // =====================
        $file = self::DEMO_DIR . 'demo-02-advanced.xlsx';
        $excel = Excel::open($file);
        $sheet = $excel->sheet('Demo3');

        $cells = $sheet->readCells();
        $this->assertCount(14, $cells);

        $sheet->setReadArea('a:f');
        $cells = $sheet->readCells();
        $this->assertCount(30, $cells);

        $sheet->setReadArea('a5:d6');
        $cells = $sheet->readCells();
        $this->assertCount(8, $cells);

        $excel = Excel::open($file);
        $sheet = $excel->sheet('Demo3');
        $rows = $sheet->readRows();
        $this->assertEquals(['A' => 'aaa', 'B' => 'bbb', 'C' => 'ccc', 'D' => 'ddd'], $rows[2]);
        $this->assertEquals(['A' => 6], $rows[6]);

        $sheet->setReadArea('a:f');
        $rows = $sheet->readRows(null, Excel::KEYS_ROW_ZERO_BASED);
        $this->assertEquals(['A' => 'aaa', 'B' => 'bbb', 'C' => 'ccc', 'D' => 'ddd', 'E' => null, 'F' => null], $rows[0]);
        $this->assertEquals(['A' => 6, 'B' => null, 'C' => null, 'D' => null, 'E' => null, 'F' => null], $rows[4]);

        $excel = Excel::open($file);
        $sheet = $excel->sheet('Demo3');
        $row = $sheet->readFirstRow();
        $this->assertEquals(['A' => 'aaa', 'B' => 'bbb', 'C' => 'ccc', 'D' => 'ddd'], $row);

        $row = $sheet->readFirstRowCells();
        $this->assertEquals(['A2' => 'aaa', 'B2' => 'bbb', 'C2' => 'ccc', 'D2' => 'ddd'], $row);

        $excel = Excel::open($file);
        $sheet = $excel->sheet('Demo3');
        $this->assertEquals(['A', 2], [$sheet->firstCol(), $sheet->firstRow()]);

        $file = self::DEMO_DIR . 'demo-00-test.xlsx';
        $excel = Excel::open($file);
        $sheet = $excel->sheet();
        $sheet->setReadArea('a:d');
        $rows = $sheet->readRows();
        $this->assertEquals(['A' => 'name', 'B' => 'birthday', 'C' => 'random_int', 'D' => null], $rows[1]);

        $excel = Excel::open($file);
        $sheet = $excel->sheet();
        $rows = $sheet->readRows(Excel::KEYS_FIRST_ROW);
        $this->assertEquals(['name' => 'James Bond', 'birthday' => -2205187200, 'random_int' => 4573], $rows[2]);

        $rows = [];
        foreach ($sheet->nextRow([], Excel::KEYS_FIRST_ROW) as $n => $rowData) {
            $rows[$n] = $rowData;
        }
        $this->assertEquals(['name' => 'James Bond', 'birthday' => -2205187200, 'random_int' => 4573], $rows[2]);
    }

}

