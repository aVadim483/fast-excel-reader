<?php

use avadim\FastExcelReader\Excel;
use PHPUnit\Framework\TestCase;

class XlsxStatTest extends TestCase
{
    public function testSheetStat()
    {
        $file = __DIR__ . '/test_files/wrong-dimension.xlsx';
        self::assertFileExists($file);
        $excel = Excel::open($file);
        $sheet = $excel->sheet();

        $stat = $sheet->stat();

        $this->assertSame(['min' => 3, 'max' => 5, 'count' => 3], $stat['rows']);
        $this->assertSame(['min' => 'C', 'max' => 'E', 'count' => 3], $stat['cols']);
        $this->assertSame(['total' => 9, 'filled' => 9], $stat['cells']);
    }

    public function testStatEmptyAndStyledCells()
    {
        // standard-file.xlsx contains a couple of styled-but-empty cells
        $file = __DIR__ . '/test_files/standard-file.xlsx';
        $excel = Excel::open($file);
        $stat = $excel->sheet()->stat();

        $this->assertSame(['min' => 1, 'max' => 29, 'count' => 29], $stat['rows']);
        $this->assertSame(['min' => 'A', 'max' => 'I', 'count' => 9], $stat['cols']);
        $this->assertSame(261, $stat['cells']['total']);
        $this->assertSame(259, $stat['cells']['filled']);

        // styles.xlsx: all cells are styled but empty -> filled == 0
        $excelS = Excel::open(__DIR__ . '/test_files/styles.xlsx');
        $statS = $excelS->sheet()->stat();
        $this->assertSame(10, $statS['cells']['total']);
        $this->assertSame(0, $statS['cells']['filled']);
    }

    /**
     * stat() rows/cols must match countActualDimension() on the same file
     */
    public function testStatMatchesActualDimension()
    {
        foreach (['wrong-dimension.xlsx', 'standard-file.xlsx', 'formulas.xlsx', 'empty.xlsx'] as $name) {
            $file = __DIR__ . '/test_files/' . $name;

            $stat = Excel::open($file)->sheet()->stat();
            $dim = Excel::open($file)->sheet()->countActualDimension();

            $this->assertSame($dim['rows'], $stat['rows'], "rows mismatch for $name");
            $this->assertSame($dim['cols'], $stat['cols'], "cols mismatch for $name");
        }
    }

    /**
     * stat().cells must match an independent count via readCells()
     * (total = all cells, filled = cells whose value is not null)
     */
    public function testStatCellsMatchReadCells()
    {
        foreach (['wrong-dimension.xlsx', 'standard-file.xlsx', 'formulas.xlsx', 'styles.xlsx', 'empty.xlsx'] as $name) {
            $file = __DIR__ . '/test_files/' . $name;

            $stat = Excel::open($file)->sheet()->stat();
            $cells = Excel::open($file)->sheet()->readCells();

            $gtTotal = count($cells);
            $gtFilled = count(array_filter($cells, static fn($v) => $v !== null));

            $this->assertSame($gtTotal, $stat['cells']['total'], "cells.total mismatch for $name");
            $this->assertSame($gtFilled, $stat['cells']['filled'], "cells.filled mismatch for $name");
        }
    }

    public function testWorkbookStat()
    {
        $file = __DIR__ . '/test_files/standard-file.xlsx';
        $excel = Excel::open($file);

        $stat = $excel->stat();

        $this->assertArrayHasKey('sheets', $stat);
        $this->assertArrayHasKey('total', $stat);

        // per-sheet entries are keyed by sheet name
        foreach ($excel->sheets() as $sheet) {
            $this->assertArrayHasKey($sheet->name(), $stat['sheets']);
        }

        $this->assertSame($excel->countSheets(), $stat['total']['sheets']);
        $this->assertSame(count($excel->visibleSheets()), $stat['total']['visible']);
        $this->assertSame(count($excel->hiddenSheets()), $stat['total']['hidden']);

        // totals equal the sum over sheets
        $sumTotal = $sumFilled = $sumRows = 0;
        foreach ($stat['sheets'] as $s) {
            $sumTotal += $s['cells']['total'];
            $sumFilled += $s['cells']['filled'];
            $sumRows += $s['rows']['count'];
        }
        $this->assertSame($sumTotal, $stat['total']['cells']['total']);
        $this->assertSame($sumFilled, $stat['total']['cells']['filled']);
        $this->assertSame($sumRows, $stat['total']['rows']);
    }
}
