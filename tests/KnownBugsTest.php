<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;

/**
 * Defects that exist in the current code base, pinned deliberately.
 *
 * These tests assert BROKEN behaviour on purpose. They exist so that a
 * refactoring cannot silently change it in either direction: neither by making
 * it worse, nor by accidentally "fixing" it without a conscious decision and a
 * CHANGELOG entry.
 *
 * When a bug here is fixed on purpose, replace the test with one asserting the
 * correct behaviour - do not just delete it.
 *
 * All of these live in methods that had zero test coverage before this suite.
 */
final class KnownBugsTest extends GuardTestCase
{
    /**
     * Sheet::readFirstRowCellsFrom() can never succeed.
     *
     * It forwards $columnKeys into readFirstRowCells(?bool $styleIdxInclude),
     * so even the default call with $columnKeys = [] is a TypeError.
     *
     * Root cause: readFirstRowCells() used to take a $columnKeys parameter. The
     * parameter was removed from the signature but the body still references
     * it - see the dangling `use (&$columnKeys)` in Sheet::readFirstRowCells(),
     * where the variable is undefined and therefore always null.
     *
     * @todo known bug - readFirstRowCellsFrom() is unusable
     *
     * @return void
     */
    public function testReadFirstRowCellsFromIsBroken(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $this->expectException(\TypeError::class);
        $this->expectExceptionMessageMatches('/readFirstRowCells\(\).*must be of type \?bool, array given/');

        $sheet->readFirstRowCellsFrom('B2:D6');
    }

    /**
     * The equivalent two-step call works, which confirms the defect is in the
     * argument forwarding of readFirstRowCellsFrom() and not in the reading.
     *
     * @return void
     */
    public function testReadFirstRowCellsWorksWhenAreaIsSetSeparately(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $result = $sheet->setReadArea('B2:D6')->readFirstRowCells();

        $this->assertNotEmpty($result);
    }

    /**
     * Sheet::rewind() is documented as an alias of reset(), but it discards the
     * caller's $columnKeys: the body does `$this->reset($columnKeys = [], ...)`,
     * an assignment rather than a pass-through.
     *
     * @todo known bug - rewind() ignores $columnKeys
     *
     * @return void
     */
    public function testRewindIgnoresColumnKeysUnlikeReset(): void
    {
        $columnKeys = ['A' => 'num', 'B' => 'hero'];

        $viaReset = iterator_to_array(
            Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->reset($columnKeys)
        );
        $viaRewind = iterator_to_array(
            Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->rewind($columnKeys)
        );

        $this->assertSame(['num', 'hero', 'C', 'D'], array_keys(reset($viaReset)));

        // Should have been the same as reset(), but the keys are not applied
        $this->assertSame(['A', 'B', 'C', 'D'], array_keys(reset($viaRewind)));
        $this->assertNotSame($viaReset, $viaRewind);
    }

    /**
     * Sheet::firstCol() ignores the column bounds of the read area.
     *
     * In nextRow() the assignment of area['first_row']/['first_col'] sits
     * between the row filter and the column filter (Sheet.php:1854-1857, the
     * column range is only checked on the next line), so first_col records the
     * first cell of the row as stored in the file, whatever the area says.
     * firstRow() is unaffected because the row filter runs earlier.
     *
     * @todo known bug - firstCol() does not respect the read area
     *
     * @return void
     */
    public function testFirstColIgnoresTheColumnBoundsOfTheReadArea(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();
        $sheet->setReadArea('C4:D9');

        $this->assertSame(4, $sheet->firstRow(), 'the row bound is honoured');

        // should have been 'C'
        $this->assertSame('B', $sheet->firstCol(), 'the column bound is not honoured');
    }

    /**
     * Restricting columns and asking for numeric column keys loses every value.
     *
     * With a read area in place, area['col_keys'] is pre-seeded with letter
     * keys, and _rowTemplate() turns that into a row template of nulls keyed by
     * letter. nextRow() then stores the real values under NUMERIC keys, so the
     * yielded row carries both. readCallback() afterwards keeps only columns
     * present in area['col_keys'] - which the numeric keys are not - so the
     * values are dropped and the null template is what survives.
     *
     * Row-based key modes are unaffected: only KEYS_COL_ZERO_BASED and
     * KEYS_COL_ONE_BASED collide with a column restriction.
     *
     * @todo known bug - column key modes are incompatible with a read area
     *
     * @dataProvider columnKeyModeProvider
     *
     * @param int $resultMode
     *
     * @return void
     */
    public function testColumnKeyModesLoseValuesWhenColumnsAreRestricted(int $resultMode): void
    {
        $rows = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->sheet()->setReadArea('C4:D8')->readRows(false, $resultMode);

        $first = reset($rows);

        // should have been [0 => 'Date', 1 => 'Color'] (or 1/2 for one-based)
        $this->assertSame(['C', 'D'], array_keys($first), 'keys stay alphabetic');
        $this->assertSame([null, null], array_values($first), 'and the values are gone');
    }

    /**
     * @return array<string, array{0: int}>
     */
    public function columnKeyModeProvider(): array
    {
        return [
            'col zero based' => [Excel::KEYS_COL_ZERO_BASED],
            'col one based' => [Excel::KEYS_COL_ONE_BASED],
        ];
    }

    /**
     * The same restriction expressed with setReadAreaColumns() behaves the same
     *
     * @return void
     */
    public function testColumnKeyModesLoseValuesWithSetReadAreaColumnsToo(): void
    {
        $rows = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->sheet()->setReadAreaColumns('C:D')->readRows(false, Excel::KEYS_COL_ZERO_BASED);

        $this->assertSame([null, null], array_values(reset($rows)));
    }

    /**
     * Without a column restriction the same mode works, which isolates the
     * defect to the interaction rather than to the key mode itself
     *
     * @return void
     */
    public function testColumnKeyModeWorksWithoutAReadArea(): void
    {
        $rows = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->sheet()->readRows(false, Excel::KEYS_COL_ZERO_BASED);

        $this->assertSame([0, 1, 2], array_keys(reset($rows)));
        $this->assertSame('Data of Sheet1', reset($rows)[0]);
    }

    /**
     * The generator itself does not lose the values - it yields the null
     * template AND the real values side by side. The loss happens later, in
     * readCallback(). Pinned separately so a fix can be verified at both levels.
     *
     * @return void
     */
    public function testGeneratorYieldsBothTemplateAndValuesUnderTheBug(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet()->setReadArea('C4:D8');

        foreach ($sheet->nextRow([], Excel::KEYS_COL_ZERO_BASED) as $row) {
            $this->assertSame(['C', 'D', 0, 1], array_keys($row));
            $this->assertSame(['Date', 'Color'], [$row[0], $row[1]]);
            break;
        }
    }

    /**
     * rewind() does forward the remaining arguments, so only $columnKeys is
     * affected. Pinning this narrows the blast radius of an eventual fix.
     *
     * @return void
     */
    public function testRewindForwardsOtherArguments(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $rows = iterator_to_array($sheet->rewind([], Excel::KEYS_ONE_BASED, null, 2));

        $this->assertCount(2, $rows);
        $this->assertSame([1, 2, 3, 4], array_keys(reset($rows)));
    }
}
