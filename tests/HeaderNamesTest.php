<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Tests\Support\XlsxBuilder;

/**
 * withHeader() with an explicit list of column names.
 *
 * The header row is consumed either way; the list only decides where the names
 * come from. Names are positional - first name to the first column of the read
 * area - so calling code needs no knowledge of column letters, which is what
 * makes the same call work on a sheet whose data does not start at A1.
 *
 * The name matches the sibling fast-excel-writer, which writes a header row
 * with writeHeader().
 */
final class HeaderNamesTest extends GuardTestCase
{
    private const XLS_DIR = __DIR__ . '/test_files/xls/';

    /**
     * @return void
     */
    public function testNamesReplaceTheHeaderRowValues(): void
    {
        $rows = Excel::open(self::fixture('demo-00-test.xlsx'))
            ->sheet()->withHeader(['num', 'hero', 'born', 'rnd'])->readRows();

        $this->assertSame(['num', 'hero', 'born', 'rnd'], array_keys(reset($rows)));
        $this->assertSame('James Bond', $rows[2]['hero']);
    }

    /**
     * The header row is still skipped, exactly as with no argument
     *
     * @return void
     */
    public function testHeaderRowIsStillConsumed(): void
    {
        $withNames = Excel::open(self::fixture('demo-00-test.xlsx'))
            ->sheet()->withHeader(['a', 'b', 'c', 'd'])->readRows();
        $withoutNames = Excel::open(self::fixture('demo-00-test.xlsx'))
            ->sheet()->withHeader()->readRows();

        $this->assertSame(array_keys($withoutNames), array_keys($withNames), 'the same rows are returned');
        $this->assertSame(
            array_values(array_values($withoutNames)[0]),
            array_values(array_values($withNames)[0]),
            'only the keys differ, never the values'
        );
    }

    /**
     * Positions are counted from the first column of the read area, so the same
     * call works on a sheet whose data starts at B2
     *
     * @return void
     */
    public function testNamesArePositionalNotLetterBased(): void
    {
        $rows = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->selectSheet('Demo1')->withHeader(['name', 'date', 'color'])->readRows();

        $first = reset($rows);

        $this->assertSame(['name', 'date', 'color'], array_keys($first));
        $this->assertSame('Giovanni', $rows[5]['name']);
        $this->assertSame('Red', $rows[5]['color']);
    }

    /**
     * Columns past the end of the list keep the name from the header row, so a
     * partial list renames only what it covers
     *
     * @return void
     */
    public function testShorterListLeavesTheRestFromTheHeaderRow(): void
    {
        $rows = Excel::open(self::fixture('demo-00-test.xlsx'))
            ->sheet()->withHeader(['num', 'hero'])->readRows();

        $this->assertSame(['num', 'hero', 'birthday', 'random_int'], array_keys(reset($rows)));
    }

    /**
     * A list longer than the sheet is not an error; the surplus is unused
     *
     * @return void
     */
    public function testLongerListIsTruncated(): void
    {
        $rows = Excel::open(self::fixture('demo-00-test.xlsx'))
            ->sheet()->withHeader(['a', 'b', 'c', 'd', 'e', 'f'])->readRows();

        $this->assertSame(['a', 'b', 'c', 'd'], array_keys(reset($rows)));
    }

    /**
     * An empty list and no argument mean the same thing
     *
     * @return void
     */
    public function testEmptyListBehavesLikeNoArgument(): void
    {
        $withEmpty = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->withHeader([])->readRows();
        $withNone = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->withHeader()->readRows();

        $this->assertSame($withNone, $withEmpty);
    }

    /**
     * The list is taken in order, whatever keys it happens to have
     *
     * @return void
     */
    public function testListKeysAreIgnored(): void
    {
        $rows = Excel::open(self::fixture('demo-00-test.xlsx'))
            ->sheet()->withHeader([7 => 'num', 3 => 'hero'])->readRows();

        $this->assertSame(['num', 'hero', 'birthday', 'random_int'], array_keys(reset($rows)));
    }

    /**
     * Works the same for XLS, since the header handling is shared code
     *
     * @return void
     */
    public function testXlsMatchesXlsx(): void
    {
        $names = ['name', 'date', 'color'];

        $xlsx = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->selectSheet('Demo1')->withHeader($names)->readRows();
        $xls = Excel::open(self::XLS_DIR . 'demo-02-advanced.xls')
            ->selectSheet('Demo1')->withHeader($names)->readRows();

        $this->assertSame($xlsx, $xls);
    }

    /**
     * And for CSV, which has its own reader
     *
     * @return void
     */
    public function testCsv(): void
    {
        $rows = Excel::openCsv(self::TEST_FILES_DIR . 'test.csv')->withHeader(['a', 'b'])->readRows();

        $first = reset($rows);

        $this->assertSame('a', array_keys($first)[0]);
        $this->assertSame('b', array_keys($first)[1]);
        $this->assertNotSame('a', array_keys($first)[2], 'the third column keeps its own header');
    }

    /**
     * Names combine with a read area: positions start at the area, not at A
     *
     * @return void
     */
    public function testWithReadArea(): void
    {
        $file = XlsxBuilder::withRows([
            1 => ['A' => 'skip', 'B' => 'h1', 'C' => 'h2'],
            2 => ['A' => 'x', 'B' => 'v1', 'C' => 'v2'],
            3 => ['A' => 'y', 'B' => 'v3', 'C' => 'v4'],
        ])->build();

        $rows = Excel::open($file)->sheet()->setReadArea('B1:C3')->withHeader(['first', 'second'])->readRows();

        $this->assertSame(['first', 'second'], array_keys(reset($rows)));
        $this->assertSame(['v1', 'v2'], array_values($rows[2]));
    }

    /**
     * setReadArea() rebuilds the area, so it must be called before withHeader()
     *
     * @return void
     */
    public function testReadAreaResetsTheNames(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $rows = $sheet->withHeader(['num', 'hero'])->setReadArea('A1:D4')->readRows();

        $this->assertSame(['A', 'B', 'C', 'D'], array_keys(reset($rows)), 'the area reset both the names and header mode');
    }
}
