<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;

/**
 * Merged cell lookups.
 *
 * getMergedCells(), isMerged() and mergedRange() had no coverage at all. They
 * are worth pinning because they read the sheet XML a second time, from a
 * different anchor: merge definitions live AFTER </sheetData>, so _readBottom()
 * scans past the data to the end and caches the result. That second pass is
 * separate from the nextRow() pass and has to keep working once the reading is
 * split across a base class and a format-specific subclass.
 */
final class SheetMergedCellsTest extends GuardTestCase
{
    /**
     * The merge map is keyed by the top-left cell of each range
     *
     * @return void
     */
    public function testMergedCellsAreKeyedByTopLeftCell(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        $this->assertSame(['B2' => 'B2:D2'], $sheet->getMergedCells());
    }

    /**
     * A sheet without merges reports an empty map, not null
     *
     * @return void
     */
    public function testSheetWithoutMergesReturnsEmptyArray(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $this->assertSame([], $sheet->getMergedCells());
    }

    /**
     * isMerged() answers for the anchor cell and for the cells it covers
     *
     * @return void
     */
    public function testIsMergedCoversTheWholeRange(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        $this->assertTrue($sheet->isMerged('B2'), 'anchor cell');
        $this->assertTrue($sheet->isMerged('C2'), 'covered cell');
        $this->assertTrue($sheet->isMerged('D2'), 'last covered cell');

        $this->assertFalse($sheet->isMerged('E2'), 'just outside the range');
        $this->assertFalse($sheet->isMerged('B4'), 'different row');
    }

    /**
     * Cell addresses are case-insensitive
     *
     * @return void
     */
    public function testIsMergedIsCaseInsensitive(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        $this->assertTrue($sheet->isMerged('b2'));
        $this->assertSame('B2:D2', $sheet->mergedRange('c2'));
    }

    /**
     * mergedRange() returns the containing range, or null outside any merge
     *
     * @return void
     */
    public function testMergedRangeResolvesTheContainingRange(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        $this->assertSame('B2:D2', $sheet->mergedRange('B2'));
        $this->assertSame('B2:D2', $sheet->mergedRange('C2'));
        $this->assertNull($sheet->mergedRange('B5'));
    }

    /**
     * A sheet with many merges resolves each of them
     *
     * @return void
     */
    public function testMultipleMergesOnOneSheet(): void
    {
        $sheet = Excel::open(self::fixture('demo-04-styles.xlsx'))->sheet();
        $merged = $sheet->getMergedCells();

        $this->assertCount(12, $merged);

        foreach ($merged as $anchor => $range) {
            $this->assertTrue($sheet->isMerged($anchor), $anchor . ' must report as merged');
            $this->assertSame($range, $sheet->mergedRange($anchor));
        }
    }

    /**
     * The lookup is cached and must survive a full data read in either order
     *
     * @return void
     */
    public function testMergeLookupIsStableAroundReads(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        $before = $sheet->getMergedCells();
        $sheet->readRows();
        $after = $sheet->getMergedCells();

        $this->assertSame($before, $after);

        $other = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();
        $other->readRows();

        $this->assertSame($before, $other->getMergedCells(), 'reading first must give the same answer');
    }

    /**
     * Merged cells carry their value in the anchor only; the covered cells read
     * back as empty. Pinned because it is the behaviour callers rely on when
     * they use the merge map to fill values forward.
     *
     * @return void
     */
    public function testOnlyTheAnchorCellCarriesTheValue(): void
    {
        $cells = Excel::open(self::fixture('demo-02-advanced.xlsx'))->readCells();
        $rows = Excel::open(self::fixture('demo-02-advanced.xlsx'))->readRows();

        $this->assertSame('Data of Sheet1', $cells['B2']);

        // the covered cells exist as styled but valueless cells, so they are
        // present in the result and read back as null - not as empty strings
        // and not absent
        $this->assertArrayHasKey('C2', $cells);
        $this->assertNull($cells['C2']);
        $this->assertNull($cells['D2']);

        $this->assertArrayHasKey('C', $rows[2]);
        $this->assertNull($rows[2]['C']);
    }
}
