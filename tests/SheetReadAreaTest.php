<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;

/**
 * Read areas, column ranges and the hidden state they share.
 *
 * The read area is not a plain value object: nextRow() takes a COPY of it
 * ($readArea = $this->area) but writes 'first_row'/'first_col' back into the
 * ORIGINAL, while _rowTemplate() mutates $this->area['col_names'] from inside
 * the deferred generator body - and readCallback() reads that very key in its
 * own loop to decide which columns survive.
 *
 * So the column filter depends on a mutation performed by the generator on its
 * first iteration. Passing the area around by reference, or rebuilding it in a
 * base class, changes the outcome. These tests pin the observable results.
 *
 * demo-02-advanced.xlsx is the fixture of choice here: its data starts at B2,
 * not A1, so the lazily computed row/column offsets actually differ.
 */
final class SheetReadAreaTest extends GuardTestCase
{
    /**
     * firstRow()/firstCol() are populated as a side effect of reading, and must
     * report the real position of the data, not the sheet origin
     *
     * @return void
     */
    public function testFirstRowAndFirstColOnOffsetSheet(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        $this->assertSame(2, $sheet->firstRow());
        $this->assertSame('B', $sheet->firstCol());
    }

    /**
     * Those accessors work without any explicit read, by triggering one.
     *
     * The triggered readFirstRow() stops as soon as the second row starts, so
     * the cost stays constant - it does not scan the sheet. demo-02-advanced
     * has no row 3, hence the counter lands on 4.
     *
     * @return void
     */
    public function testFirstRowIsResolvedLazilyWithoutExplicitRead(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        // no read call before this line
        $this->assertSame(2, $sheet->firstRow());
        $this->assertSame(4, $sheet->getReadRowNum(), 'reading stops on the row after the first one');
    }

    /**
     * Guard against readFirstRow() degrading into a full scan
     *
     * @return void
     */
    public function testFirstRowDoesNotScanTheWholeSheet(): void
    {
        $sheet = Excel::open(self::fixture('demo-100k-rows.xlsx'))->sheet();

        $sheet->readFirstRow();

        $this->assertSame(2, $sheet->getReadRowNum(), 'a 100k row sheet must not be walked to find row 1');
    }

    /**
     * Narrowing the area moves the reported first cell in both directions
     *
     * @return void
     */
    public function testFirstRowFollowsTheReadArea(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();
        $sheet->setReadArea('C4:D9');

        $this->assertSame(4, $sheet->firstRow());
        $this->assertSame('C', $sheet->firstCol());
    }

    /**
     * setReadArea() is sticky: it mutates the sheet rather than returning a view
     *
     * @return void
     */
    public function testSetReadAreaIsSticky(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $narrowed = $sheet->setReadArea('B2:C4')->readRows();
        $again = $sheet->readRows();

        $this->assertSame($narrowed, $again, 'the area must still apply to the next read');
        $this->assertSame([2, 3, 4], array_keys($again));
        $this->assertSame(['B', 'C'], array_keys($again[2]));
    }

    /**
     * The *From() helpers are setReadArea() plus a read, so they are sticky too
     *
     * @return void
     */
    public function testReadRowsFromIsStickyAsWell(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $from = $sheet->readRowsFrom('B2:C4');

        $this->assertSame($from, $sheet->readRows());
    }

    /**
     * setReadAreaColumns() restricts columns while keeping every row
     *
     * @return void
     */
    public function testSetReadAreaColumnsKeepsAllRows(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $rows = $sheet->setReadAreaColumns('B:D')->readRows();

        $this->assertCount(527, $rows);
        $this->assertSame(['B', 'C', 'D'], array_keys(reset($rows)));
    }

    /**
     * The column filter in readCallback() consults area['col_names'], which
     * _rowTemplate() fills in from inside the generator. Combining a column
     * range with an explicit key map exercises exactly that path.
     *
     * @return void
     */
    public function testColumnRangeCombinedWithColumnKeys(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $rows = $sheet->setReadAreaColumns('B:D')->readRows(['B' => 'beta', 'D' => 'delta']);

        $first = reset($rows);

        $this->assertSame(['beta', 'C', 'delta'], array_keys($first), 'renamed and untouched columns must coexist');
    }

    /**
     * Same interaction, but with the keys taken from the first row
     *
     * @return void
     */
    public function testColumnRangeWithFirstRowKeys(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $rows = $sheet->setReadAreaColumns('B:D', true)->readRows();
        $first = reset($rows);

        $this->assertCount(3, $first);
        $this->assertNotSame(['B', 'C', 'D'], array_keys($first), 'header values must replace the letters');
    }

    /**
     * withHeader() is the declarative form of the same thing
     *
     * @return void
     */
    public function testWithHeaderMatchesFirstRowKeys(): void
    {
        $viaHeader = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->withHeader()->readRows();
        $viaFlag = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->readRows(true);

        $this->assertSame($viaFlag, $viaHeader);
        $this->assertSame(['#', 'name', 'birthday', 'random_int'], array_keys(reset($viaHeader)));
    }

    /**
     * from() sets an open-ended area anchored at a cell
     *
     * @return void
     */
    public function testFromAnchorsTheAreaWithoutUpperBound(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        $rows = $sheet->from('C3')->readRows();

        // row 3 does not exist in this fixture, so reading starts at row 4
        $this->assertSame(4, array_key_first($rows));
        $this->assertSame(11, array_key_last($rows));
        $this->assertSame(['C', 'D'], array_keys(reset($rows)));
    }

    /**
     * A row key mode applied on top of a narrowed area: the offset is computed
     * from the first row that passes the filter, not from the sheet origin
     *
     * @return void
     */
    public function testRowKeyOffsetIsRelativeToTheFilteredFirstRow(): void
    {
        $zeroBased = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->sheet()->setReadArea('C3:D8')->readRows(false, Excel::KEYS_ROW_ZERO_BASED);

        // rows 4..8 - the fixture has no row 3 - renumbered from zero
        $this->assertSame([0, 1, 2, 3, 4], array_keys($zeroBased));
        $this->assertSame(['Date', 'Color'], array_values(reset($zeroBased)));
    }

    /**
     * An area that contains no data yields an empty result rather than throwing
     *
     * @return void
     */
    public function testAreaOutsideDataYieldsNothing(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $this->assertSame([], $sheet->setReadArea('Z100:AA200')->readRows());
    }

    /**
     * A single cell is treated as the top-left corner of an open-ended area,
     * the same way from() behaves - not as a one-cell range
     *
     * @return void
     */
    public function testSingleCellAreaIsOpenEnded(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $rows = $sheet->setReadArea('B2')->readRows();

        $this->assertSame([2, 3, 4], array_keys($rows));
        $this->assertSame(['B', 'C', 'D'], array_keys($rows[2]));
        $this->assertSame('James Bond', $rows[2]['B']);
    }

    /**
     * Selecting another sheet must hand out an independent read area.
     *
     * getSheetNames() is keyed by sheet id, not by position, so the names have
     * to be taken through array_values().
     *
     * @return void
     */
    public function testReadAreaIsPerSheet(): void
    {
        $excel = Excel::open(self::fixture('demo-00-test.xlsx'));
        $names = array_values($excel->getSheetNames());

        $narrowed = $excel->selectSheet($names[0])->setReadArea('A1:B2')->readRows();
        $this->assertSame([1, 2], array_keys($narrowed));
        $this->assertSame(['A', 'B'], array_keys($narrowed[1]));

        $other = $excel->selectSheet($names[1])->readRows();

        $this->assertCount(11, $other, 'the second sheet must not inherit the two-row area');
    }
}
