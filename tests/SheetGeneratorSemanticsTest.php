<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;

/**
 * Generator lifecycle and the ORDER of side effects.
 *
 * Snapshots compare return values, so they cannot detect that work moved from
 * "first iteration" to "call time". That distinction matters here: nextRow()
 * contains a yield, so its entire body - dimension(), the recursive
 * readFirstRow() for KEYS_FIRST_ROW, opening the sheet XML, resetting the row
 * counters - is deferred until the generator is first advanced. reset() relies
 * on exactly that, zeroing countReadRows itself AFTER creating the generator.
 *
 * Splitting nextRow() into a plain wrapper plus an inner generator - the
 * obvious move when extracting a base class - would change this silently.
 */
final class SheetGeneratorSemanticsTest extends GuardTestCase
{
    /**
     * Creating the generator must not read anything yet
     *
     * @return void
     */
    public function testNextRowIsLazyUntilFirstIteration(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $generator = $sheet->nextRow();

        $this->assertSame(0, $sheet->getReadRowNum(), 'nextRow() must not read before the generator is advanced');

        $generator->current();

        $this->assertSame(1, $sheet->getReadRowNum(), 'the first row must be read on the first advance');
    }

    /**
     * The same laziness must hold for the KEYS_FIRST_ROW path, which triggers a
     * recursive readFirstRow() inside the generator body
     *
     * @return void
     */
    public function testNextRowWithFirstRowKeysIsAlsoLazy(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $generator = $sheet->nextRow([], Excel::KEYS_FIRST_ROW);

        $this->assertSame(0, $sheet->getReadRowNum());

        $row = $generator->current();

        $this->assertSame(['#', 'name', 'birthday', 'random_int'], array_keys($row));
    }

    /**
     * The sharpest probe of the deferred body: reset() does NOT clear
     * readRowNum. The counter keeps the stale value from the previous read and
     * is only rewritten once the generator is actually advanced, because the
     * assignment lives inside the generator body, after the point where
     * execution suspends.
     *
     * If nextRow() is ever split into a plain wrapper plus an inner generator,
     * the "after reset" value below turns into 0 and this test fires.
     *
     * @return void
     */
    public function testResetLeavesRowNumStaleUntilGeneratorIsAdvanced(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $this->assertSame(0, $sheet->getReadRowNum(), 'fresh sheet');

        $sheet->readRows();
        $this->assertSame(4, $sheet->getReadRowNum(), 'after a full read');

        $generator = $sheet->reset();
        $this->assertSame(4, $sheet->getReadRowNum(), 'reset() must not touch the counter by itself');

        $generator->current();
        $this->assertSame(1, $sheet->getReadRowNum(), 'the counter is rewritten inside the generator body');
    }

    /**
     * readNextRow() walks the sheet one row at a time and stops at the end
     *
     * @return void
     */
    public function testReadNextRowWalksEveryRowThenReturnsFalsy(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $rows = [];
        $rowNums = [];
        while ($row = $sheet->readNextRow()) {
            $rows[] = $row;
            $rowNums[] = $sheet->getReadRowNum();
        }

        $this->assertCount(4, $rows);
        $this->assertSame([1, 2, 3, 4], $rowNums);
        $this->assertSame($sheet->readRows()[1], $rows[0]);
    }

    /**
     * readNextRow() auto-resets on first use, so it needs no explicit reset()
     *
     * @return void
     */
    public function testReadNextRowAutoResets(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $first = $sheet->readNextRow();

        $this->assertNotEmpty($first);
        $this->assertSame(1, $sheet->getReadRowNum());
    }

    /**
     * An explicit reset() mid-iteration restarts the walk from the top
     *
     * @return void
     */
    public function testResetRestartsIteration(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $firstPass = [$sheet->readNextRow(), $sheet->readNextRow()];

        $sheet->reset();
        $secondPass = [$sheet->readNextRow(), $sheet->readNextRow()];

        $this->assertSame($firstPass, $secondPass);
    }

    /**
     * A full read is repeatable: the sheet must not be left in a spent state
     *
     * @return void
     */
    public function testReadRowsIsRepeatable(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        $this->assertSame($sheet->readRows(), $sheet->readRows());
        $this->assertSame($sheet->readCells(), $sheet->readCells());
    }

    /**
     * Mixing the generator API with the array API on one sheet must not make
     * either of them lose rows
     *
     * @return void
     */
    public function testGeneratorAndArrayApiDoNotInterfere(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();
        $expected = $sheet->readRows();

        $sheet->readNextRow();
        $sheet->readNextRow();

        $this->assertSame($expected, $sheet->readRows(), 'a partial generator walk must not affect readRows()');

        $sheet->reset();
        $walked = [];
        while ($row = $sheet->readNextRow()) {
            $walked[] = $row;
        }

        $this->assertSame(array_values($expected), $walked);
    }

    /**
     * $rowLimit caps the number of yielded rows
     *
     * @return void
     */
    public function testRowLimitCapsIteration(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $this->assertCount(1, iterator_to_array($sheet->nextRow([], null, null, 1)));
        $this->assertCount(5, iterator_to_array($sheet->nextRow([], null, null, 5)));

        // 0 means "no limit"
        $this->assertCount(527, iterator_to_array($sheet->nextRow([], null, null, 0)));
    }

    /**
     * getReadRowNum() reports the sheet row number, not the yielded key, and so
     * is unaffected by the KEYS_* offsets
     *
     * @return void
     */
    public function testGetReadRowNumIsIndependentOfKeyMode(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        foreach ($sheet->nextRow([], Excel::KEYS_ROW_ZERO_BASED) as $key => $row) {
            // data starts at row 2, so the zero-based key trails the real number
            $this->assertSame(2, $sheet->getReadRowNum());
            $this->assertSame(0, $key);
            break;
        }
    }
}
