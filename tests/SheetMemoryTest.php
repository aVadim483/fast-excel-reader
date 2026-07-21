<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;

/**
 * Streaming guarantees.
 *
 * The whole point of the library is that a sheet is walked node by node and
 * never materialised. A refactoring can preserve every returned value and still
 * destroy that property - for example by buffering rows inside an intermediate
 * generator, or by collecting cells into an array before yielding.
 *
 * The thresholds below are deliberately loose: they are there to catch a
 * change in ORDER OF MAGNITUDE, not to police a few kilobytes.
 *
 * These are the only slow tests in the suite - they walk 100k rows several
 * times. Run them with Xdebug disabled (XDEBUG_MODE=off), which cuts the time
 * from roughly two minutes to under six seconds.
 */
final class SheetMemoryTest extends GuardTestCase
{
    private const HUGE_FILE = 'demo-100k-rows.xlsx';

    /**
     * Walking 100k rows with the generator must not accumulate them
     *
     * @return void
     */
    public function testNextRowDoesNotAccumulateRows(): void
    {
        $sheet = Excel::open(self::fixture(self::HUGE_FILE))->sheet();

        $before = memory_get_usage();
        $count = 0;
        $last = null;
        foreach ($sheet->nextRow() as $row) {
            $last = $row;
            $count++;
        }
        $growth = memory_get_usage() - $before;

        $this->assertSame(100000, $count);
        $this->assertNotEmpty($last);
        $this->assertLessThan(
            2 * 1024 * 1024,
            $growth,
            sprintf('walking %d rows grew the heap by %d bytes - the reader stopped streaming', $count, $growth)
        );
    }

    /**
     * Per-row memory must stay flat: the cost of the last row must not exceed
     * the cost of the first one
     *
     * @return void
     */
    public function testMemoryStaysFlatAcrossTheSheet(): void
    {
        $sheet = Excel::open(self::fixture(self::HUGE_FILE))->sheet();

        $samples = [];
        $index = 0;
        foreach ($sheet->nextRow() as $row) {
            if ($index === 100 || $index === 50000 || $index === 99999) {
                $samples[$index] = memory_get_usage();
            }
            $index++;
        }

        $this->assertCount(3, $samples);

        $drift = $samples[99999] - $samples[100];
        $this->assertLessThan(
            1024 * 1024,
            $drift,
            sprintf('heap drifted by %d bytes between row 100 and row 100000', $drift)
        );
    }

    /**
     * readNextRow() is the same walk driven manually and must stream too
     *
     * @return void
     */
    public function testReadNextRowStreams(): void
    {
        $sheet = Excel::open(self::fixture(self::HUGE_FILE))->sheet();

        $before = memory_get_usage();
        $count = 0;
        while ($sheet->readNextRow()) {
            $count++;
            if ($count >= 50000) {
                break;
            }
        }
        $growth = memory_get_usage() - $before;

        $this->assertSame(50000, $count);
        $this->assertLessThan(2 * 1024 * 1024, $growth, sprintf('grew by %d bytes', $growth));
    }

    /**
     * Opening the workbook must not read the sheet data. This is what makes
     * selecting one sheet out of many cheap.
     *
     * @return void
     */
    public function testOpeningAWorkbookDoesNotLoadSheetData(): void
    {
        $before = memory_get_usage();

        $excel = Excel::open(self::fixture(self::HUGE_FILE));
        $sheet = $excel->sheet();

        $growth = memory_get_usage() - $before;

        $this->assertNotNull($sheet);
        $this->assertLessThan(
            4 * 1024 * 1024,
            $growth,
            sprintf('opening a 100k row workbook allocated %d bytes', $growth)
        );
    }

    /**
     * A row limit must stop the walk early rather than read everything and
     * slice afterwards
     *
     * @return void
     */
    public function testRowLimitStopsEarly(): void
    {
        $sheet = Excel::open(self::fixture(self::HUGE_FILE))->sheet();

        $started = microtime(true);
        $rows = iterator_to_array($sheet->nextRow([], null, null, 10));
        $elapsed = microtime(true) - $started;

        $this->assertCount(10, $rows);
        $this->assertLessThan(
            1.0,
            $elapsed,
            sprintf('reading 10 rows out of 100000 took %.2fs - the limit is not short-circuiting', $elapsed)
        );
    }
}
