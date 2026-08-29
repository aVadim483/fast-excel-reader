<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Tests\Support\XlsxBuilder;

/**
 * Characters that XML cannot carry as is are stored as "_xHHHH_" (the ST_Xstring type),
 * a literal "_xHHHH_" of the source text is stored as "_x005F_xHHHH_".
 *
 * The reader must undo both: the first form back to the character, the second one back
 * to the literal text. Only the second form used to be handled, so a value written with
 * a CR in it - by Excel itself or by the sibling fast-excel-writer - came back as the
 * literal "_x000D_" instead of the character.
 */
final class XlsxEscapedCharsTest extends GuardTestCase
{
    public function testEscapedCharactersAreDecoded(): void
    {
        $file = XlsxBuilder::withRows([
            1 => [
                'A' => 'line 1_x000D_line 2',
                'B' => '_x0000__x001F__x007F_',
                'C' => '_x0041_',
                'D' => 'no escapes here',
            ],
        ])->build();

        $result = Excel::open($file)->readCells();

        $this->assertSame("line 1\rline 2", $result['A1']);
        $this->assertSame("\x00\x1F\x7F", $result['B1']);
        $this->assertSame('A', $result['C1']);
        $this->assertSame('no escapes here', $result['D1']);
    }

    /**
     * "_x005F_xHHHH_" carries a literal "_xHHHH_" of the source text and must not be
     * decoded twice - the restored literal stays as it is.
     *
     * @return void
     */
    public function testEscapedLiteralSequencesAreRestored(): void
    {
        $file = XlsxBuilder::withRows([
            1 => [
                'A' => '_x005F_x000D_',
                'B' => '_x005F_x005F_x0041_',
                'C' => '_x005F_x0041__x000D_',
            ],
        ])->build();

        $result = Excel::open($file)->readCells();

        $this->assertSame('_x000D_', $result['A1']);
        $this->assertSame('_x005F_x0041_', $result['B1']);
        $this->assertSame("_x0041_\r", $result['C1']);
    }

    /**
     * A formula caches its string result in the cell as t="str", and that result carries
     * the same escapes - a concatenation producing a line break is an everyday case.
     *
     * @return void
     */
    public function testFormulaStringResultIsDecoded(): void
    {
        $file = XlsxBuilder::withRows([
            1 => [
                'A' => 'line 1_x000D_line 2',
                'B' => '_x005F_x000D_',
            ],
        ])->withFormulas([
            'A1' => 'B1&amp;C1',
            'B1' => 'C1',
        ])->build();

        $result = Excel::open($file)->readCells();

        $this->assertSame("line 1\rline 2", $result['A1']);
        $this->assertSame('_x000D_', $result['B1']);
    }

    /**
     * Decoding must not change the type of a numeric formula result: "str" stays its own
     * type, so a calculated number is still cast to int/float rather than returned as text.
     *
     * @return void
     */
    public function testNumericFormulaResultKeepsItsType(): void
    {
        $file = XlsxBuilder::withRows([
            1 => [
                'A' => 2,
                'B' => 1.5,
                'C' => '2',
            ],
        ])->withFormulas([
            'A1' => '1+1',
            'B1' => '3/2',
            'C1' => '1+1',
        ])->build();

        $result = Excel::open($file)->readCells();

        $this->assertSame(2, $result['A1']);
        $this->assertSame(1.5, $result['B1']);
        // Written as t="str", yet numeric - the reader casts it as it did before
        $this->assertSame(2, $result['C1']);
    }

    /**
     * The same decoding applies to values taken from the shared string table
     *
     * @return void
     */
    public function testSharedStringsAreDecoded(): void
    {
        $file = self::fixture('escaped-chars.xlsx');

        $result = Excel::open($file)->readCells();

        $this->assertSame("line 1\rline 2", $result['A1']);
        $this->assertSame('_x000D_', $result['A2']);
    }
}
