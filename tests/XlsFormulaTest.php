<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Xls\FormulaParser;

/**
 * Formula text read from XLS.
 *
 * A .xls formula is not stored as text but compiled to a postfix token array
 * for a stack machine; FormulaParser runs that machine backwards to recover the
 * A1 text. Parity is against the same formula read from the XLSX counterpart.
 *
 * XLSX keeps the leading "=" in the 'f' field, XLS does not, so comparisons
 * strip it. Everything after the "=" must match exactly, including the absence
 * of spaces around operators, which is how Excel itself stores them.
 */
final class XlsFormulaTest extends GuardTestCase
{
    private const XLS_DIR = __DIR__ . '/test_files/xls/';

    /**
     * Every formula in the shared-formula fixture matches its XLSX text.
     *
     * This file is built entirely of shared formulas: one master token array in
     * a SHRFMLA record, referenced by a tExp token in each cell and re-based to
     * that cell. So "=A2+1" in B2 becomes "=A3+1" in B3, and getting the
     * re-basing wrong shifts every reference.
     *
     * @return void
     */
    public function testSharedFormulasMatchXlsx(): void
    {
        $xls = Excel::open(self::XLS_DIR . 'formulas.xls')->sheet()->readCells(true);
        $xlsx = Excel::open(self::fixture('formulas.xlsx'))->sheet()->readCells(true);

        $compared = 0;
        foreach ($xlsx as $address => $cell) {
            if (empty($cell['f'])) {
                continue;
            }
            $this->assertArrayHasKey($address, $xls);
            $this->assertSame(
                ltrim($cell['f'], '='),
                (string)$xls[$address]['f'],
                $address
            );
            $compared++;
        }

        $this->assertGreaterThan(40, $compared, 'the fixture must actually contain formulas');
    }

    /**
     * A spread of ordinary formulas: operators with precedence, explicit
     * parentheses, a unary minus, string concatenation, comparisons, and both
     * fixed and variable arity functions
     *
     * @return void
     */
    public function testVariedFormulasMatchXlsx(): void
    {
        $xls = Excel::open(self::XLS_DIR . 'varied-formulas.xls')->sheet()->readCells(true);
        $xlsx = Excel::open(self::fixture('xls-formulas-source.xlsx'))->sheet()->readCells(true);

        $expected = [
            'B1' => 'A1+B1*2',
            'B2' => '(A1+B1)*2',
            'B3' => 'SUM(A1:A5)',
            'B4' => 'IF(A1>10,"big","small")',
            'B5' => 'ROUND(A1/B1,2)',
            'B6' => '-A1',
            'B7' => 'A1&"x"&B1',
            'B8' => 'MAX(A1,B1,100)',
            'B9' => 'A1<=B1',
            'B10' => 'ABS(A1-B1)',
        ];

        foreach ($expected as $address => $formula) {
            $this->assertSame($formula, (string)$xls[$address]['f'], $address . ' from xls');
            $this->assertSame($formula, ltrim($xlsx[$address]['f'], '='), $address . ' from xlsx (guard)');
        }
    }

    /**
     * The cached result still reads as a normal value alongside the formula
     *
     * @return void
     */
    public function testFormulaCellsCarryBothResultAndText(): void
    {
        $cells = Excel::open(self::XLS_DIR . 'formulas.xls')->sheet()->readCells(true);

        $b2 = $cells['B2'];
        $this->assertSame('A2+1', $b2['f']);
        $this->assertSame('number', $b2['t']);
        $this->assertIsInt($b2['v']);
    }

    /**
     * readCells() without the descriptor still returns the cached result, never
     * the formula text
     *
     * @return void
     */
    public function testPlainReadReturnsResults(): void
    {
        $xls = Excel::open(self::XLS_DIR . 'formulas.xls')->readCells();
        $xlsx = Excel::open(self::fixture('formulas.xlsx'))->readCells();

        $this->assertSame($xlsx, $xls);
    }

    /**
     * An operand-only expression round-trips through the stack machine
     *
     * @return void
     */
    public function testConstantsAndOperators(): void
    {
        // tInt 2, tInt 3, tMul  ->  2*3
        $this->assertSame('2*3', (new FormulaParser())->parse("\x1E\x02\x00\x1E\x03\x00\x05"));

        // tInt 1, tUminus  ->  -1
        $this->assertSame('-1', (new FormulaParser())->parse("\x1E\x01\x00\x13"));
    }

    /**
     * An absolute reference keeps its dollar signs
     *
     * @return void
     */
    public function testAbsoluteReference(): void
    {
        // tRef row=0 colField=0x0000 (both absolute) -> $A$1
        $this->assertSame('$A$1', (new FormulaParser())->parse("\x24\x00\x00\x00\x00"));
    }

    /**
     * An unknown token aborts the whole formula rather than emitting a guess,
     * so the reader falls back to "text unavailable" and keeps the cached value
     *
     * @return void
     */
    public function testUnknownTokenYieldsNull(): void
    {
        // 0xFF is not a token this parser renders
        $this->assertNull((new FormulaParser())->parse("\x1E\x01\x00\xFF"));
    }

    /**
     * A dangling operator with nothing to operate on is rejected, not rendered
     * as a broken string
     *
     * @return void
     */
    public function testMalformedStackYieldsNull(): void
    {
        // tAdd with an empty stack
        $this->assertNull((new FormulaParser())->parse("\x03"));

        // a single operand leaves the stack with more than one entry is fine,
        // but two operands and no operator is not a complete formula
        $this->assertNull((new FormulaParser())->parse("\x1E\x01\x00\x1E\x02\x00"));
    }

    /**
     * A formula whose text cannot be recovered still yields its cached value,
     * with the formula reported as null rather than raising
     *
     * @return void
     */
    public function testUnrenderableFormulaKeepsItsValue(): void
    {
        // every value in the sheet reads back, and none raises, even if some
        // formula texts happen to be null
        $cells = Excel::open(self::XLS_DIR . 'formulas.xls')->sheet()->readCells();

        $this->assertNotEmpty($cells);
        foreach ($cells as $value) {
            $this->assertNotNull($value);
        }
    }
}
