<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Tests\Support\XlsxBuilder;

/**
 * Regression tests for defects found while building the refactoring safety net.
 *
 * Each of these lived in a method that had no test coverage at all, which is
 * why they went unnoticed. They were first pinned as broken behaviour, then
 * fixed; the assertions below describe the corrected behaviour and exist to
 * keep it that way.
 */
final class FixedDefectsTest extends GuardTestCase
{
    /**
     * readFirstRowCellsFrom() used to throw unconditionally: it forwarded
     * $columnKeys into readFirstRowCells(?bool $styleIdxInclude), so even a
     * default call was a TypeError. The result is keyed by cell address, so -
     * like readCellsFrom() - the method takes no column keys at all.
     *
     * @return void
     */
    public function testReadFirstRowCellsFromReturnsTheFirstRow(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $result = $sheet->readFirstRowCellsFrom('B2:D6');

        $this->assertSame(['B2', 'C2', 'D2'], array_keys($result));
        $this->assertSame('Pasta', $result['B2']);
    }

    /**
     * The same call must agree with setting the area separately
     *
     * @return void
     */
    public function testReadFirstRowCellsFromMatchesTheTwoStepForm(): void
    {
        $inOneStep = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet()->readFirstRowCellsFrom('B2:D6');
        $inTwoSteps = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet()->setReadArea('B2:D6')->readFirstRowCells();

        $this->assertSame($inTwoSteps, $inOneStep);
    }

    /**
     * The style-index variant still works through the second parameter
     *
     * @return void
     */
    public function testReadFirstRowCellsFromWithStyleIndexes(): void
    {
        $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

        $result = $sheet->readFirstRowCellsFrom('B2:D6', true);

        $this->assertSame(['v', 's', 'f', 't', 'o'], array_keys($result['B2']));
        $this->assertSame('Pasta', $result['B2']['v']);
    }

    /**
     * Restricting columns and asking for numeric column keys used to return
     * nothing but nulls: values were stored under numeric keys while the row
     * template held letters, and readCallback() then dropped the numeric half.
     *
     * @dataProvider columnKeyModeProvider
     *
     * @param int $resultMode
     * @param array $expectedKeys
     *
     * @return void
     */
    public function testColumnKeyModesKeepValuesWhenColumnsAreRestricted(int $resultMode, array $expectedKeys): void
    {
        $rows = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->sheet()->setReadArea('C4:D8')->readRows(false, $resultMode);

        $first = reset($rows);

        $this->assertSame($expectedKeys, array_keys($first));
        $this->assertSame(['Date', 'Color'], array_values($first));
    }

    /**
     * @return array<string, array{0: int, 1: array}>
     */
    public function columnKeyModeProvider(): array
    {
        return [
            'col zero based' => [Excel::KEYS_COL_ZERO_BASED, [0, 1]],
            'col one based' => [Excel::KEYS_COL_ONE_BASED, [1, 2]],
        ];
    }

    /**
     * setReadAreaColumns() is affected by the same code path
     *
     * @return void
     */
    public function testColumnKeyModesKeepValuesWithSetReadAreaColumns(): void
    {
        $rows = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->sheet()->setReadAreaColumns('C:D')->readRows(false, Excel::KEYS_COL_ZERO_BASED);

        $this->assertSame([0, 1], array_keys($rows[4]));
        $this->assertSame(['Date', 'Color'], array_values($rows[4]));
        $this->assertSame([18316800, 'Red'], array_values($rows[5]));
    }

    /**
     * Behaviour without a read area is unchanged
     *
     * @return void
     */
    public function testColumnKeyModeStillWorksWithoutAReadArea(): void
    {
        $rows = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->sheet()->readRows(false, Excel::KEYS_COL_ZERO_BASED);

        $this->assertSame([0, 1, 2], array_keys(reset($rows)));
        $this->assertSame('Data of Sheet1', reset($rows)[0]);
    }

    /**
     * The generator no longer yields the letter template alongside the numeric
     * values - the row carries one set of keys, not two
     *
     * @return void
     */
    public function testGeneratorYieldsASingleSetOfKeys(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet()->setReadArea('C4:D8');

        foreach ($sheet->nextRow([], Excel::KEYS_COL_ZERO_BASED) as $row) {
            $this->assertSame([0, 1], array_keys($row));
            $this->assertSame(['Date', 'Color'], [$row[0], $row[1]]);
            break;
        }
    }

    /**
     * An explicit column name must still win over the numeric key, otherwise
     * the rename requested by the caller would be lost. Guards the interaction
     * that the fix above had to preserve.
     *
     * @return void
     */
    public function testExplicitColumnKeysWinOverNumericColumnKeys(): void
    {
        $rows = Excel::open(self::fixture('demo-02-advanced.xlsx'))
            ->sheet()->setReadArea('B4:D8')
            ->readRows(['C' => 'when'], Excel::KEYS_COL_ONE_BASED);

        $first = reset($rows);

        $this->assertSame([1, 'when', 3], array_keys($first), 'the named column keeps its name, the others are numeric');
        $this->assertSame('Date', $first['when']);
        $this->assertSame('Name', $first[1]);
        $this->assertSame('Color', $first[3]);
    }

    /**
     * rewind() is documented as an alias of reset(), but it used to assign to
     * its own parameter instead of forwarding it, silently dropping the keys
     *
     * @return void
     */
    public function testRewindForwardsColumnKeysLikeReset(): void
    {
        $columnKeys = ['A' => 'num', 'B' => 'hero'];

        $viaReset = iterator_to_array(
            Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->reset($columnKeys)
        );
        $viaRewind = iterator_to_array(
            Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->rewind($columnKeys)
        );

        $this->assertSame(['num', 'hero', 'C', 'D'], array_keys(reset($viaRewind)));
        $this->assertSame($viaReset, $viaRewind);
    }

    /**
     * The remaining arguments were always forwarded and still are
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

    /**
     * firstCol() used to report the first cell of the row as stored in the
     * file, ignoring the column bounds of the read area, because first_col was
     * recorded between the row filter and the column filter
     *
     * @return void
     */
    public function testFirstColRespectsTheColumnBoundsOfTheReadArea(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();
        $sheet->setReadArea('C4:D9');

        $this->assertSame(4, $sheet->firstRow());
        $this->assertSame('C', $sheet->firstCol());
    }

    /**
     * Without an area the reported first cell is the first one that exists
     *
     * @return void
     */
    public function testFirstColWithoutAReadArea(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

        $this->assertSame(2, $sheet->firstRow());
        $this->assertSame('B', $sheet->firstCol());
    }

    /**
     * A column-only restriction is honoured as well
     *
     * @return void
     */
    public function testFirstColWithColumnRangeOnly(): void
    {
        $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();
        $sheet->setReadAreaColumns('C:D');

        $this->assertSame('C', $sheet->firstCol());
    }

    /**
     * readCellsWithStylesFrom() used to return bare values: it called
     * readCells() instead of readCellsWithStyles(), so it dropped the styles
     * its own name promises and passed the style key into a bool parameter.
     *
     * @return void
     */
    public function testReadCellsWithStylesFromReturnsStyles(): void
    {
        $inOneStep = Excel::open(self::fixture('demo-04-styles.xlsx'))->sheet()->readCellsWithStylesFrom('A1:B2');
        $inTwoSteps = Excel::open(self::fixture('demo-04-styles.xlsx'))->sheet()->setReadArea('A1:B2')->readCellsWithStyles();

        $this->assertSame($inTwoSteps, $inOneStep);

        $this->assertSame(['v', 's', 'f', 't', 'o'], array_keys($inOneStep['A2']));
        $this->assertSame('This is demo XLSX-sheet', $inOneStep['A2']['v']);
        $this->assertArrayHasKey('font', $inOneStep['A2']['s']);
    }

    /**
     * readCellsWithStyles($styleKey) is documented with 'fill-color' as its
     * example, yet that never worked: the key was looked up on the nested
     * style, where properties sit one level down inside their group, so the
     * whole style was returned instead.
     *
     * @return void
     */
    public function testReadCellsWithStylesNarrowsToASingleProperty(): void
    {
        $cells = Excel::open(self::fixture('demo-04-styles.xlsx'))
            ->sheet()->setReadArea('A1:E12')->readCellsWithStyles('fill-color');

        $this->assertSame(['fill-color' => '#9FC63C'], $cells['A1']['s']);
    }

    /**
     * A group name still selects the whole group, and an unknown key falls back
     * to the complete style rather than losing it
     *
     * @return void
     */
    public function testReadCellsWithStylesKeyFallbacks(): void
    {
        $sheet = Excel::open(self::fixture('demo-04-styles.xlsx'))->sheet();

        $byGroup = $sheet->setReadArea('A1:B2')->readCellsWithStyles('font');
        $this->assertArrayHasKey('font-name', $byGroup['A1']['s']['font']);

        $unknown = Excel::open(self::fixture('demo-04-styles.xlsx'))
            ->sheet()->setReadArea('A1:B2')->readCellsWithStyles('no-such-property');
        $this->assertArrayHasKey('font', $unknown['A1']['s']);
    }

    /**
     * The complete styles were parsed by walking childNodes of <fonts>, <fills>,
     * <borders> and <cellXfs>, which in a pretty-printed styles.xml also holds
     * the whitespace between the tags. Every writer that indents its output -
     * and several do - made getCompleteStyleByIdx() and readCellsWithStyles()
     * die with "Call to undefined method DOMText::getAttribute()".
     *
     * Only element children are styles now. Indentation must therefore make no
     * difference at all, which is exactly what this compares.
     *
     * @return void
     */
    public function testCompleteStylesIgnoreIndentationInStylesXml(): void
    {
        $rows = [1 => ['A' => 'first'], 2 => ['A' => 2]];

        $indented = XlsxBuilder::withRows($rows)->withStyles(self::indentedStyles())->build();
        // same document, only the whitespace between the tags is gone
        $compact = XlsxBuilder::withRows($rows)
            ->withStyles((string)preg_replace('/>\s+</', '><', self::indentedStyles()))
            ->build();

        $fromIndented = Excel::open($indented)->getCompleteStyleByIdx(1);
        $fromCompact = Excel::open($compact)->getCompleteStyleByIdx(1);

        $this->assertSame($fromCompact, $fromIndented);

        // and the second entry of every table is really the one that was picked,
        // i.e. the text nodes did not shift the indexes either
        $this->assertSame('14', $fromIndented['font']['font-size']);
        $this->assertSame(1, $fromIndented['font']['font-style-bold']);
        $this->assertSame('#00FF00', $fromIndented['fill']['fill-color']);
        $this->assertSame('thin', $fromIndented['border']['border-left-style']);
        $this->assertSame('0.000', $fromIndented['format']['format-pattern']);
    }

    /**
     * A styles.xml with two entries in every table, written the way an indenting
     * writer would write it
     *
     * @return string
     */
    private static function indentedStyles(): string
    {
        return <<<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="0.000"/>
  </numFmts>
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FFFF0000"/>
      <name val="Arial"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF00FF00"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="2">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left style="thin"/>
      <right style="thin"/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="1" borderId="1" xfId="0" applyFont="1">
      <alignment horizontal="center"/>
    </xf>
  </cellXfs>
</styleSheet>
XML;
    }
}
