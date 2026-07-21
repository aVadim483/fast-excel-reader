<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Xls\XlsBook;
use avadim\FastExcelReader\Xls\XlsSheet;

/**
 * Reading XLS (BIFF8) workbooks.
 *
 * The strongest statement this suite can make is parity: the .xls fixtures were
 * converted by LibreOffice from the .xlsx demo files, so for every sheet whose
 * content survived the conversion unchanged, both readers must return exactly
 * the same array - same values, same types, same keys.
 *
 * Two sheets of demo-00-test are deliberately excluded from the parity check.
 * There, values that Excel had stored as dates were written out as literal text
 * by the converter, so the files genuinely differ; that is a property of the
 * fixtures, not of the readers.
 */
final class XlsReaderTest extends GuardTestCase
{
    private const XLS_DIR = __DIR__ . '/test_files/xls/';

    /**
     * @return void
     */
    public function testOpenDetectsXlsBySignature(): void
    {
        $book = Excel::open(self::XLS_DIR . 'demo-00-test.xls');

        $this->assertInstanceOf(XlsBook::class, $book);
        $this->assertInstanceOf(XlsSheet::class, $book->sheet());
    }

    /**
     * The extension is not consulted, only the signature
     *
     * @return void
     */
    public function testIsXls(): void
    {
        $this->assertTrue(Excel::isXls(self::XLS_DIR . 'demo-00-test.xls'));
        $this->assertFalse(Excel::isXls(self::fixture('demo-00-test.xlsx')));
        $this->assertFalse(Excel::isXls(self::XLS_DIR . 'no-such-file.xls'));
    }

    /**
     * @return void
     */
    public function testOpenXlsIsExplicit(): void
    {
        $book = Excel::openXls(self::XLS_DIR . 'demo-00-test.xls');

        $this->assertInstanceOf(XlsBook::class, $book);
        $this->assertSame(['Sheet1', 'Sheet2', 'Sheet3'], array_values($book->getSheetNames()));
    }

    /**
     * @return void
     */
    public function testOpeningAnXlsxAsXlsThrows(): void
    {
        $this->expectException(Exception::class);

        Excel::openXls(self::fixture('demo-00-test.xlsx'));
    }

    /**
     * Every sheet of every fixture must match its XLSX counterpart exactly.
     *
     * @dataProvider parityProvider
     *
     * @param string $xlsxFile
     * @param string $xlsFile
     * @param string $sheetName
     *
     * @return void
     */
    public function testReadRowsMatchXlsx(string $xlsxFile, string $xlsFile, string $sheetName): void
    {
        $xlsx = Excel::open(self::fixture($xlsxFile))->selectSheet($sheetName)->readRows();
        $xls = Excel::open(self::XLS_DIR . $xlsFile)->selectSheet($sheetName)->readRows();

        $this->assertSame($xlsx, $xls);
    }

    /**
     * @dataProvider parityProvider
     *
     * @param string $xlsxFile
     * @param string $xlsFile
     * @param string $sheetName
     *
     * @return void
     */
    public function testReadCellsMatchXlsx(string $xlsxFile, string $xlsFile, string $sheetName): void
    {
        $xlsx = Excel::open(self::fixture($xlsxFile))->selectSheet($sheetName)->readCells();
        $xls = Excel::open(self::XLS_DIR . $xlsFile)->selectSheet($sheetName)->readCells();

        $this->assertSame($xlsx, $xls);
    }

    /**
     * The key modes live in AbstractSheet, so this also proves the XLS reader
     * really reuses the shared implementation rather than a copy of it
     *
     * @dataProvider parityProvider
     *
     * @param string $xlsxFile
     * @param string $xlsFile
     * @param string $sheetName
     *
     * @return void
     */
    public function testKeyModesMatchXlsx(string $xlsxFile, string $xlsFile, string $sheetName): void
    {
        $modes = [
            Excel::KEYS_ORIGINAL,
            Excel::KEYS_ZERO_BASED,
            Excel::KEYS_ONE_BASED,
            Excel::KEYS_RELATIVE,
            Excel::KEYS_FIRST_ROW,
            Excel::KEYS_SWAP,
        ];

        foreach ($modes as $mode) {
            $xlsx = Excel::open(self::fixture($xlsxFile))->selectSheet($sheetName)->readRows(false, $mode);
            $xls = Excel::open(self::XLS_DIR . $xlsFile)->selectSheet($sheetName)->readRows(false, $mode);

            $this->assertSame($xlsx, $xls, 'result mode ' . $mode);
        }
    }

    /**
     * Read areas are shared code too
     *
     * @return void
     */
    public function testReadAreaMatchesXlsx(): void
    {
        $xlsx = Excel::open(self::fixture('demo-02-advanced.xlsx'))->selectSheet('Demo1')->readRowsFrom('C4:D8');
        $xls = Excel::open(self::XLS_DIR . 'demo-02-advanced.xls')->selectSheet('Demo1')->readRowsFrom('C4:D8');

        $this->assertSame($xlsx, $xls);
        $this->assertNotEmpty($xls);
    }

    /**
     * @return array<string, array{0: string, 1: string, 2: string}>
     */
    public function parityProvider(): array
    {
        return [
            // demo-00-test Sheet1 and Sheet3 are excluded: the converter turned
            // date cells into literal text, so the files themselves differ
            'demo-00-test Sheet2' => ['demo-00-test.xlsx', 'demo-00-test.xls', 'Sheet2'],
            'demo-02-advanced Demo1' => ['demo-02-advanced.xlsx', 'demo-02-advanced.xls', 'Demo1'],
            'demo-02-advanced Demo2' => ['demo-02-advanced.xlsx', 'demo-02-advanced.xls', 'Demo2'],
            'demo-02-advanced Demo3' => ['demo-02-advanced.xlsx', 'demo-02-advanced.xls', 'Demo3'],
            'demo-05-datetime Sheet1' => ['demo-05-datetime.xlsx', 'demo-05-datetime.xls', 'Sheet1'],
        ];
    }

    /**
     * Dates are the part most likely to differ silently: a serial number only
     * becomes a date because its number format says so
     *
     * @return void
     */
    public function testDatesMatchXlsxUnderEveryFormatterSetting(): void
    {
        $variants = [
            'default' => static function ($book) {
            },
            'formatted' => static function ($book) {
                $book->setDateFormat('Y-m-d H:i:s');
            },
            'raw timestamps' => static function ($book) {
                $book->dateFormatter(false);
            },
        ];

        foreach ($variants as $label => $configure) {
            $xlsx = Excel::open(self::fixture('demo-05-datetime.xlsx'));
            $xls = Excel::open(self::XLS_DIR . 'demo-05-datetime.xls');
            $configure($xlsx);
            $configure($xls);

            $this->assertSame($xlsx->readRows(), $xls->readRows(), $label);
        }
    }

    /**
     * dateFormatter(null) asks for the untouched serial number, and there the
     * two formats cannot agree on representation: XLSX hands back the literal
     * text from the XML, while XLS only ever held a binary double, so the
     * original digits do not exist in the file. The numeric value is the same.
     *
     * @return void
     */
    public function testUnformattedDateSerialsAgreeNumerically(): void
    {
        $xlsx = Excel::open(self::fixture('demo-05-datetime.xlsx'));
        $xls = Excel::open(self::XLS_DIR . 'demo-05-datetime.xls');
        $xlsx->dateFormatter(null);
        $xls->dateFormatter(null);

        $fromXlsx = $xlsx->readRows();
        $fromXls = $xls->readRows();

        $this->assertSame(array_keys($fromXlsx), array_keys($fromXls));

        foreach ($fromXlsx as $rowNum => $row) {
            $this->assertIsString($row['B'], 'XLSX keeps the serial as written');
            $this->assertIsFloat($fromXls[$rowNum]['B'], 'XLS only has the double');
            $this->assertEqualsWithDelta((float)$row['B'], $fromXls[$rowNum]['B'], 1e-9, 'row ' . $rowNum);
        }
    }

    /**
     * @return void
     */
    public function testSheetMetadata(): void
    {
        $book = Excel::open(self::XLS_DIR . 'demo-02-advanced.xls');

        $this->assertSame(3, $book->countSheets());
        $this->assertTrue($book->sheetExists('Demo2'));

        $sheet = $book->selectSheet('Demo1');

        $this->assertSame('Demo1', $sheet->name());
        $this->assertSame('B2:D11', $sheet->dimension());
        $this->assertSame(2, $sheet->firstRow());
        $this->assertSame('B', $sheet->firstCol());
        $this->assertTrue($sheet->isVisible());

        // the offset of the sheet substream stands in for the inner path
        $this->assertNotSame('', $sheet->path());
    }

    /**
     * @return void
     */
    public function testMergedCellsMatchXlsx(): void
    {
        $xlsx = Excel::open(self::fixture('demo-02-advanced.xlsx'))->selectSheet('Demo1');
        $xls = Excel::open(self::XLS_DIR . 'demo-02-advanced.xls')->selectSheet('Demo1');

        $this->assertSame(['B2' => 'B2:D2'], $xls->getMergedCells());
        $this->assertSame($xlsx->getMergedCells(), $xls->getMergedCells());

        $this->assertTrue($xls->isMerged('C2'));
        $this->assertSame('B2:D2', $xls->mergedRange('C2'));
        $this->assertFalse($xls->isMerged('B5'));
    }

    /**
     * RK is a compressed number: two flag bits are stolen from the low end, so
     * both the integer and the truncated-double encodings have to be handled,
     * with and without the division by 100
     *
     * @return void
     */
    public function testNumbersOfEveryShape(): void
    {
        $rows = Excel::open(self::XLS_DIR . 'demo-02-advanced.xls')->selectSheet('Demo2')->readRows();

        $found = ['int' => false, 'float' => false, 'negative' => false];
        foreach ($rows as $row) {
            foreach ($row as $value) {
                if (is_int($value)) {
                    $found['int'] = true;
                    if ($value < 0) {
                        $found['negative'] = true;
                    }
                }
                elseif (is_float($value)) {
                    $found['float'] = true;
                }
            }
        }

        $this->assertTrue($found['int'], 'integers must be read as integers');
        $this->assertTrue($found['float'], 'fractional numbers must stay floats');
    }

    /**
     * A whole number must not come back as a float, which is what the XLSX
     * reader does and what assertSame in the parity tests depends on
     *
     * @return void
     */
    public function testWholeNumbersAreIntegers(): void
    {
        $cells = Excel::open(self::XLS_DIR . 'demo-00-test.xls')->selectSheet('Sheet1')->readCells(true);

        $this->assertSame('number', $cells['A2']['t']);
        $this->assertIsInt($cells['A2']['v']);
        $this->assertSame(1, $cells['A2']['v']);
    }

    /**
     * Cell descriptors have the same shape as in the XLSX reader
     *
     * @return void
     */
    public function testCellDescriptorShape(): void
    {
        $cells = Excel::open(self::XLS_DIR . 'demo-00-test.xls')->selectSheet('Sheet1')->readCells(true);

        $this->assertSame(['v', 's', 'f', 't', 'o'], array_keys($cells['B2']));
        $this->assertSame('James Bond', $cells['B2']['v']);
        $this->assertSame('string', $cells['B2']['t']);
        $this->assertIsInt($cells['B2']['s']);
    }

    /**
     * Sheets are reached by seeking to the offset stored in BOUNDSHEET, so
     * selecting them out of order must work and must not disturb each other
     *
     * @return void
     */
    public function testSheetsAreIndependent(): void
    {
        $book = Excel::open(self::XLS_DIR . 'demo-02-advanced.xls');

        $third = $book->selectSheet('Demo3')->readRows();
        $first = $book->selectSheet('Demo1')->readRows();
        $firstAgain = $book->selectSheet('Demo1')->readRows();
        $thirdAgain = $book->selectSheet('Demo3')->readRows();

        $this->assertSame($first, $firstAgain);
        $this->assertSame($third, $thirdAgain);
        $this->assertNotSame($first, $third);
    }

    /**
     * Two sheets read at the same time must not share a read position
     *
     * @return void
     */
    public function testInterleavedGeneratorsDoNotInterfere(): void
    {
        $book = Excel::open(self::XLS_DIR . 'demo-02-advanced.xls');
        $expectedFirst = $book->selectSheet('Demo1')->readRows();

        $one = $book->getSheet('Demo1')->nextRow();
        $two = $book->getSheet('Demo2')->nextRow();

        $collected = [];
        $one->current();
        $two->current();
        foreach ($one as $rowNum => $row) {
            $collected[$rowNum] = $row;
            $two->next();
        }

        $this->assertSame($expectedFirst, $collected);
    }

    /**
     * The generator must not accumulate rows
     *
     * @return void
     */
    public function testReadingStreams(): void
    {
        $sheet = Excel::open(self::XLS_DIR . 'continue-sst.xls')->sheet();

        $before = memory_get_usage();
        $count = 0;
        foreach ($sheet->nextRow() as $row) {
            $count++;
        }
        $growth = memory_get_usage() - $before;

        $this->assertSame(220, $count);
        $this->assertLessThan(2 * 1024 * 1024, $growth, sprintf('grew by %d bytes', $growth));
    }

    /**
     * Strings that were split across CONTINUE boundaries in the shared string
     * table must arrive at the cells intact
     *
     * @return void
     */
    public function testValuesFromASplitSharedStringTable(): void
    {
        $rows = Excel::open(self::XLS_DIR . 'continue-sst.xls')->readRows();

        $this->assertCount(220, $rows);

        foreach ($rows as $rowNum => $row) {
            $this->assertMatchesRegularExpression('/^ascii-\d{4}-x{34}$/', $row['A'], 'row ' . $rowNum);
            $this->assertMatchesRegularExpression('/^кир-\d{4}-я{28}$/u', $row['B'], 'row ' . $rowNum);
        }
    }

    /**
     * @return void
     */
    public function testRowLimit(): void
    {
        $sheet = Excel::open(self::XLS_DIR . 'continue-sst.xls')->sheet();

        $this->assertCount(1, iterator_to_array($sheet->nextRow([], null, null, 1)));
        $this->assertCount(7, iterator_to_array($sheet->nextRow([], null, null, 7)));
    }

    /**
     * @return void
     */
    public function testReadNextRowAndReset(): void
    {
        $sheet = Excel::open(self::XLS_DIR . 'demo-00-test.xls')->selectSheet('Sheet1');

        $rows = [];
        while ($row = $sheet->readNextRow()) {
            $rows[] = $row;
        }

        $this->assertCount(4, $rows);
        $this->assertSame(array_values($sheet->readRows()), $rows);
    }
}
