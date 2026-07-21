<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Xls\XlsStyle;

/**
 * Cell formatting read from XLS.
 *
 * BIFF stores formatting differently from OOXML: fonts sit in their own
 * records, but the fill and all four borders are bit fields inside the XF
 * record itself, and colours are indices into a palette rather than literal
 * values. All of that is rebuilt into the same tables the XLSX reader
 * produces, so that a style is the same array whichever format it came from.
 *
 * Parity here is per property rather than per cell. The converter renumbered
 * the number formats - "General" became custom format 164 instead of builtin 0
 * - so format-num-id cannot agree, and on merged ranges it propagated the
 * anchor's format across the whole range where Excel had not.
 */
final class XlsStyleTest extends GuardTestCase
{
    private const XLS_DIR = __DIR__ . '/test_files/xls/';
    private const AREA = 'A1:E12';

    /**
     * Properties that must be identical in both formats for every cell.
     *
     * Colours are the interesting ones: XLSX writes them literally, XLS stores
     * a palette index, so agreement means the PALETTE record was applied right.
     *
     * @return void
     */
    public function testStylePropertiesMatchXlsx(): void
    {
        $properties = [
            'font-name',
            'font-family',
            'font-charset',
            'fill-pattern',
            'fill-color',
            'format-pattern',
            'format-category',
            'border-left-style',
            'border-right-style',
        ];

        $xlsx = Excel::open(self::fixture('demo-04-styles.xlsx'))->sheet()->setReadArea(self::AREA)->readCellStyles(true);
        $xls = Excel::open(self::XLS_DIR . 'demo-04-styles.xls')->sheet()->setReadArea(self::AREA)->readCellStyles(true);

        $this->assertNotEmpty($xlsx);
        $this->assertSame(array_keys($xlsx), array_keys($xls));

        foreach ($properties as $property) {
            $fromXlsx = [];
            $fromXls = [];
            foreach ($xlsx as $cell => $style) {
                $fromXlsx[$cell] = $style[$property] ?? null;
                $fromXls[$cell] = $xls[$cell][$property] ?? null;
            }

            $this->assertSame($fromXlsx, $fromXls, $property);
        }
    }

    /**
     * At least one cell must actually carry each of them, otherwise the test
     * above would pass on a sheet with no formatting at all
     *
     * @return void
     */
    public function testTheComparisonIsNotVacuous(): void
    {
        $styles = Excel::open(self::XLS_DIR . 'demo-04-styles.xls')->sheet()->setReadArea(self::AREA)->readCellStyles(true);

        $seen = [];
        foreach ($styles as $style) {
            foreach ($style as $key => $value) {
                if ($value !== null && $value !== '') {
                    $seen[$key] = true;
                }
            }
        }

        foreach (['font-name', 'font-size', 'font-style-bold', 'fill-pattern', 'fill-color', 'border-left-style'] as $key) {
            $this->assertArrayHasKey($key, $seen, $key . ' must occur somewhere in the fixture');
        }
    }

    /**
     * Fill colours come from the PALETTE record, which replaces colour indices
     * 8 and up
     *
     * @return void
     */
    public function testPaletteColours(): void
    {
        $styles = Excel::open(self::XLS_DIR . 'demo-04-styles.xls')->sheet()->setReadArea(self::AREA)->readCellStyles(true);

        $this->assertSame('solid', $styles['A1']['fill-pattern']);
        $this->assertSame('#9FC63C', $styles['A1']['fill-color']);
        $this->assertSame('#EEEEEE', $styles['A6']['fill-color']);

        foreach ($styles as $cell => $style) {
            if (isset($style['fill-color'])) {
                $this->assertMatchesRegularExpression('/^#[0-9A-F]{6}$/', $style['fill-color'], $cell);
            }
        }
    }

    /**
     * Font records are indexed with a hole: index 4 does not exist, so every
     * index above it refers to one record earlier than it appears to. Getting
     * that wrong shifts every font from the fifth onwards.
     *
     * @return void
     */
    public function testFontsResolveThroughTheIndexHole(): void
    {
        $styles = Excel::open(self::XLS_DIR . 'demo-04-styles.xls')->sheet()->setReadArea(self::AREA)->readCellStyles(true);

        // the title carries the largest font in the sheet
        $this->assertSame('24', $styles['A2']['font-size']);
        $this->assertSame(1, $styles['A2']['font-style-bold']);
        $this->assertSame('Arial', $styles['A2']['font-name']);

        // an ordinary cell keeps the default one
        $this->assertSame('10', $styles['A1']['font-size']);
        $this->assertArrayNotHasKey('font-style-bold', array_filter($styles['A1'], static function ($v) {
            return $v !== null;
        }));
    }

    /**
     * Borders are packed into the XF record as nibbles for the style and 7 bit
     * fields for the colour, so a wrong offset shows up as the wrong edge
     *
     * @return void
     */
    public function testBordersAreDecodedPerEdge(): void
    {
        $styles = Excel::open(self::XLS_DIR . 'demo-04-styles.xls')->sheet()->setReadArea(self::AREA)->readCellStyles(true);

        // top left corner of a bordered table: thick on the outside, thin inside
        $this->assertSame('thick', $styles['A6']['border-left-style']);
        $this->assertSame('thick', $styles['A6']['border-top-style']);
        $this->assertSame('thin', $styles['A6']['border-right-style']);

        // one column in, the left edge is an inner one
        $this->assertSame('thin', $styles['B6']['border-left-style']);
        $this->assertSame('thick', $styles['B6']['border-top-style']);

        $this->assertSame('#000000', $styles['A6']['border-left-color']);
    }

    /**
     * Alignment lives in a single byte of the XF record
     *
     * @return void
     */
    public function testAlignment(): void
    {
        $styles = Excel::open(self::XLS_DIR . 'demo-04-styles.xls')->sheet()->setReadArea(self::AREA)->readCellStyles(true);

        $this->assertSame('center', $styles['A6']['format-align-horizontal']);
        $this->assertSame('center', $styles['A6']['format-align-vertical']);

        // the defaults, general and bottom, are left out rather than spelled out
        $this->assertArrayNotHasKey('format-align-horizontal', array_filter($styles['A1'], static function ($v) {
            return $v !== null;
        }));
    }

    /**
     * The nested form groups the same data by kind
     *
     * @return void
     */
    public function testNestedStyleShape(): void
    {
        $styles = Excel::open(self::XLS_DIR . 'demo-04-styles.xls')->sheet()->setReadArea('A1:B2')->readCellStyles();

        $this->assertSame(['format', 'font', 'fill', 'border'], array_keys($styles['A1']));
        $this->assertSame('Arial', $styles['A1']['font']['font-name']);
        $this->assertSame('solid', $styles['A1']['fill']['fill-pattern']);
    }

    /**
     * getCompleteStyleByIdx() is shared code; it must compose XLS tables too
     *
     * @return void
     */
    public function testCompleteStyleByIndex(): void
    {
        $book = Excel::open(self::XLS_DIR . 'demo-04-styles.xls');
        $cells = $book->sheet()->setReadArea(self::AREA)->readCells(true);

        $styleIdx = $cells['A1']['s'];
        $nested = $book->getCompleteStyleByIdx($styleIdx);
        $flat = $book->getCompleteStyleByIdx($styleIdx, true);

        $this->assertArrayHasKey('font', $nested);
        $this->assertArrayHasKey('fill', $nested);
        $this->assertSame('#9FC63C', $flat['fill-color']);
    }

    /**
     * readCellsWithStyles() combines values and styles, and is shared code
     *
     * @return void
     */
    public function testReadCellsWithStyles(): void
    {
        $cells = Excel::open(self::XLS_DIR . 'demo-04-styles.xls')->sheet()->readCellsWithStylesFrom('A1:B2');

        $this->assertArrayHasKey('A2', $cells);
        $this->assertSame('This is demo XLSX-sheet', $cells['A2']['v']);
        $this->assertSame('24', $cells['A2']['s']['font']['font-size']);
    }

    /**
     * A single style property can be requested on its own
     *
     * @return void
     */
    public function testReadCellsWithSingleStyleKey(): void
    {
        $cells = Excel::open(self::XLS_DIR . 'demo-04-styles.xls')
            ->sheet()->setReadArea(self::AREA)->readCellsWithStyles('fill-color');

        $this->assertSame(['fill-color' => '#9FC63C'], $cells['A1']['s']);
    }

    /**
     * Number formats still drive value typing after the style tables were added
     *
     * @return void
     */
    public function testNumberFormatsStillTypeValues(): void
    {
        $xlsx = Excel::open(self::fixture('demo-05-datetime.xlsx'))->readRows();
        $xls = Excel::open(self::XLS_DIR . 'demo-05-datetime.xls')->readRows();

        $this->assertSame($xlsx, $xls);
    }

    /**
     * The lookup tables are the whole of the mapping between the two formats,
     * so a wrong entry silently renames a style
     *
     * @return void
     */
    public function testLookupTables(): void
    {
        $this->assertNull(XlsStyle::BORDER_STYLES[0], 'no border');
        $this->assertSame('thin', XlsStyle::BORDER_STYLES[1]);
        $this->assertSame('thick', XlsStyle::BORDER_STYLES[5]);
        $this->assertSame('double', XlsStyle::BORDER_STYLES[6]);

        $this->assertSame('none', XlsStyle::FILL_PATTERNS[0]);
        $this->assertSame('solid', XlsStyle::FILL_PATTERNS[1]);

        $this->assertNull(XlsStyle::ALIGN_HORIZONTAL[0], 'general is the default and is omitted');
        $this->assertSame('center', XlsStyle::ALIGN_HORIZONTAL[2]);
        $this->assertNull(XlsStyle::ALIGN_VERTICAL[2], 'bottom is the default and is omitted');
    }

    /**
     * A workbook with no PALETTE record must still resolve colours, from the
     * standard indexed table
     *
     * @return void
     */
    public function testColoursWithoutAPaletteRecord(): void
    {
        $styles = Excel::open(self::XLS_DIR . 'demo-00-test.xls')->sheet()->readCellStyles(true);

        $this->assertNotEmpty($styles);
        foreach ($styles as $cell => $style) {
            foreach (['fill-color', 'font-color', 'border-left-color'] as $key) {
                if (isset($style[$key])) {
                    $this->assertMatchesRegularExpression('/^#[0-9A-F]{6}$/', $style[$key], $cell . ' ' . $key);
                }
            }
        }
    }
}
