<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Tests\Support\XlsxBuilder;

/**
 * RESULT_MODE_ROW, TRIM_STRINGS and TREAT_EMPTY_STRING_AS_EMPTY_CELL.
 *
 * None of these flags had a single mention in the test suite before, yet all of
 * their handling sits inside nextRow() - the exact method a base-class
 * extraction has to take apart. They are asserted here directly rather than
 * only through snapshots, so that a failure names the broken flag.
 */
final class SheetResultModesTest extends GuardTestCase
{
    /**
     * RESULT_MODE_ROW wraps each row into __cells plus the raw <row> attributes
     *
     * @return void
     */
    public function testResultModeRowWrapsCellsAndRowAttributes(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $rows = iterator_to_array($sheet->nextRow([], Excel::RESULT_MODE_ROW));
        $first = reset($rows);

        $this->assertSame(['__cells', '__row'], array_keys($first));
        $this->assertSame(['A', 'B', 'C', 'D'], array_keys($first['__cells']));
        $this->assertSame('#', $first['__cells']['A']);
        $this->assertArrayHasKey('r', $first['__row'], 'the row number attribute must be exposed');
        $this->assertSame('1', $first['__row']['r']);
    }

    /**
     * The wrapper composes with the key modes
     *
     * @return void
     */
    public function testResultModeRowCombinesWithKeyModes(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $rows = iterator_to_array($sheet->nextRow([], Excel::RESULT_MODE_ROW | Excel::KEYS_ONE_BASED));
        $first = reset($rows);

        $this->assertSame([1, 2, 3, 4], array_keys($first['__cells']));
    }

    /**
     * readRows() unwraps the envelope again, so the flag is a no-op there
     *
     * @return void
     */
    public function testResultModeRowIsTransparentForReadRows(): void
    {
        $plain = Excel::open(self::fixture('demo-00-test.xlsx'))->readRows();
        $wrapped = Excel::open(self::fixture('demo-00-test.xlsx'))->readRows([], Excel::RESULT_MODE_ROW);

        $this->assertSame($plain, $wrapped);
    }

    /**
     * TRIM_STRINGS strips surrounding whitespace from string values only
     *
     * @return void
     */
    public function testTrimStringsStripsWhitespace(): void
    {
        $file = XlsxBuilder::withRows([
            1 => ['A' => '  padded  ', 'B' => 'clean', 'C' => "\ttabbed\t", 'D' => 42],
        ])->build();

        $untrimmed = Excel::open($file)->readRows();
        $trimmed = Excel::open($file)->readRows([], Excel::TRIM_STRINGS);

        $this->assertSame(['A' => '  padded  ', 'B' => 'clean', 'C' => "\ttabbed\t", 'D' => 42], $untrimmed[1]);
        $this->assertSame(['A' => 'padded', 'B' => 'clean', 'C' => 'tabbed', 'D' => 42], $trimmed[1]);
    }

    /**
     * A whitespace-only cell collapses to an empty string, it is not removed
     *
     * @return void
     */
    public function testTrimStringsTurnsWhitespaceOnlyCellsIntoEmptyStrings(): void
    {
        $file = XlsxBuilder::withRows([1 => ['A' => '   ', 'B' => 'x']])->build();

        $rows = Excel::open($file)->readRows([], Excel::TRIM_STRINGS);

        $this->assertSame(['A' => '', 'B' => 'x'], $rows[1]);
    }

    /**
     * TREAT_EMPTY_STRING_AS_EMPTY_CELL drops empty strings from the result
     *
     * @return void
     */
    public function testEmptyStringsAreDroppedWhenRequested(): void
    {
        $file = XlsxBuilder::withRows([1 => ['A' => '', 'B' => 'x', 'C' => 0]])->build();

        $kept = Excel::open($file)->readRows();
        $dropped = Excel::open($file)->readRows([], Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL);

        $this->assertSame('', $kept[1]['A']);
        $this->assertArrayNotHasKey('A', $dropped[1], 'without a read area there is no row template to fall back to');
        $this->assertSame('x', $dropped[1]['B']);
        $this->assertSame(0, $dropped[1]['C'], 'a numeric zero is not an empty string');
    }

    /**
     * With a read area in place the row template does exist, so the dropped
     * cell surfaces as null rather than disappearing. The distinction changes
     * the shape of the result and is easy to lose in a refactoring.
     *
     * @return void
     */
    public function testDroppedEmptyStringFallsBackToTemplateWhenAreaIsSet(): void
    {
        $file = XlsxBuilder::withRows([1 => ['A' => '', 'B' => 'x']])->build();

        $rows = Excel::open($file)->sheet()
            ->setReadArea('A1:B1')
            ->readRows([], Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL);

        $this->assertArrayHasKey('A', $rows[1]);
        $this->assertNull($rows[1]['A']);
    }

    /**
     * Trimming happens before the empty-string check, so the two flags compose:
     * a whitespace-only cell is dropped only when both are set
     *
     * @return void
     */
    public function testTrimAndEmptyStringFlagsCompose(): void
    {
        $file = XlsxBuilder::withRows([1 => ['A' => '   ', 'B' => 'x']])->build();

        $trimOnly = Excel::open($file)->readRows([], Excel::TRIM_STRINGS);
        $both = Excel::open($file)->readRows([], Excel::TRIM_STRINGS | Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL);

        $this->assertSame('', $trimOnly[1]['A']);
        $this->assertArrayNotHasKey('A', $both[1], 'trimming runs first, so the cell then counts as empty');
        $this->assertSame('x', $both[1]['B']);
    }

    /**
     * The flags must not disturb the key modes they are combined with
     *
     * @return void
     */
    public function testFlagsComposeWithKeyModes(): void
    {
        $file = XlsxBuilder::withRows([
            1 => ['A' => ' h1 ', 'B' => ' h2 '],
            2 => ['A' => ' v1 ', 'B' => ' v2 '],
        ])->build();

        $rows = Excel::open($file)->readRows(true, Excel::TRIM_STRINGS | Excel::KEYS_ROW_ZERO_BASED);

        $this->assertSame(['h1' => 'v1', 'h2' => 'v2'], $rows[0]);
    }

    /**
     * styleIdxInclude returns the full cell descriptor instead of a bare value.
     * TRIM_STRINGS does not reach into that structure.
     *
     * @return void
     */
    public function testStyleIdxIncludeReturnsCellDescriptors(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $rows = $sheet->readRows([], null, true);
        $cell = $rows[1]['A'];

        $this->assertSame(['v', 's', 'f', 't', 'o'], array_keys($cell));
        $this->assertSame('#', $cell['v']);
        $this->assertSame('string', $cell['t']);
    }
}
