<?php

namespace avadim\FastExcelReader\Xls;

use avadim\FastExcelHelper\Helper;

/**
 * Decompile a BIFF parsed-expression (the Ptg token array) back into A1 text
 *
 * A .xls formula is not stored as text but compiled to a postfix (reverse
 * Polish) token stream for a little stack machine. Rendering it back means
 * running that machine: operands push their text, operators and functions pop
 * their arguments and push the combined text.
 *
 * The guiding rule is to never guess. Any token this does not understand aborts
 * the whole formula and parse() returns null, so the cached result of the
 * formula stays usable and one exotic token never corrupts a file. That is why
 * the caller treats a null formula as "text unavailable", not as an error.
 */
class FormulaParser
{
    private string $data;

    private int $pos = 0;

    private int $baseRow;

    private int $baseCol;

    /** @var string[] operand stack, each entry already rendered to text */
    private array $stack = [];

    private bool $failed = false;

    /**
     * @param int $baseRow zero-based row of the cell owning the formula
     * @param int $baseCol zero-based column of the cell owning the formula
     */
    public function __construct(int $baseRow = 0, int $baseCol = 0)
    {
        $this->baseRow = $baseRow;
        $this->baseCol = $baseCol;
    }

    /**
     * Render a token array to an A1 formula string, or null if any token is
     * unsupported
     *
     * @param string $tokens The Ptg byte array
     *
     * @return string|null Without the leading "="
     */
    public function parse(string $tokens): ?string
    {
        $this->data = $tokens;
        $this->pos = 0;
        $this->stack = [];
        $this->failed = false;

        $length = strlen($tokens);
        while ($this->pos < $length && !$this->failed) {
            $this->step();
        }

        if ($this->failed || count($this->stack) !== 1) {
            return null;
        }

        return $this->stack[0];
    }

    /**
     * Consume and act on one token
     *
     * @return void
     */
    private function step(): void
    {
        $ptg = $this->byte();
        // the class bits (reference/value/array) do not change how a token is
        // rendered, so normalise them away
        $base = $ptg & 0x1F;
        if ($ptg >= 0x20) {
            $ptg = $base | 0x20;
        }

        switch ($ptg) {
            case 0x01: // tExp - a shared/array formula, handled via SHRFMLA elsewhere
            case 0x02: // tTbl
                $this->failed = true;
                break;

            // Excel stores formulas without spaces around operators, so match it
            case 0x03: $this->binary('+'); break;  // tAdd
            case 0x04: $this->binary('-'); break;  // tSub
            case 0x05: $this->binary('*'); break;  // tMul
            case 0x06: $this->binary('/'); break;  // tDiv
            case 0x07: $this->binary('^'); break;  // tPower
            case 0x08: $this->binary('&'); break;  // tConcat
            case 0x09: $this->binary('<'); break;  // tLT
            case 0x0A: $this->binary('<='); break;
            case 0x0B: $this->binary('='); break;
            case 0x0C: $this->binary('>='); break;
            case 0x0D: $this->binary('>'); break;
            case 0x0E: $this->binary('<>'); break;
            case 0x0F: $this->binary(' '); break;  // tIsect (space is the operator)
            case 0x10: $this->binary(','); break;  // tUnion
            case 0x11: $this->binary(':'); break;  // tRange

            case 0x12: $this->unaryPrefix('+'); break;  // tUplus
            case 0x13: $this->unaryPrefix('-'); break;  // tUminus
            case 0x14: $this->unarySuffix('%'); break;  // tPercent
            case 0x15: $this->parenthesis(); break;     // tParen

            case 0x16: $this->push('""'); break;       // tMissArg
            case 0x17: $this->pushString(); break;     // tStr
            case 0x1C: $this->pushError(); break;      // tErr
            case 0x1D: $this->pushBool(); break;       // tBool
            case 0x1E: $this->push((string)$this->uint16()); break; // tInt
            case 0x1F: $this->pushNumber(); break;     // tNum

            case 0x18: // tExtended, an escape byte for later token variants
                $this->failed = true;
                break;
            case 0x19: $this->attr(); break;           // tAttr

            case 0x20:
            case 0x21: $this->func(true); break;       // tFunc (fixed arg count)
            case 0x22: $this->func(false); break;      // tFuncVar (variable arg count)
            case 0x23: $this->pushName(); break;       // tName

            case 0x24: $this->pushRef(false); break;   // tRef (absolute-capable)
            case 0x25: $this->pushArea(false); break;  // tArea
            case 0x2C: $this->pushRef(true); break;    // tRefN (relative to the cell)
            case 0x2D: $this->pushArea(true); break;   // tAreaN

            default:
                // anything else - 3D refs, arrays, memory tokens, add-in calls -
                // is not rendered; give up rather than emit something wrong
                $this->failed = true;
        }
    }

    /**
     * tAttr carries optimiser hints; only the SUM shorthand affects the text
     *
     * @return void
     */
    private function attr(): void
    {
        $flags = $this->byte();
        $word = $this->uint16();

        if ($flags & 0x10) {
            // bitFuncSum: the sole argument on the stack is summed
            $arg = $this->pop();
            $this->push('SUM(' . $arg . ')');
        }
        // bitSpace (0x40), bitIf/bitChoose jump tables and bitGoto add no text
    }

    /**
     * @param bool $relative
     *
     * @return void
     */
    private function pushRef(bool $relative): void
    {
        $row = $this->uint16();
        $colField = $this->uint16();
        $this->push($this->cellRef($row, $colField, $relative));
    }

    /**
     * @param bool $relative
     *
     * @return void
     */
    private function pushArea(bool $relative): void
    {
        $rowFirst = $this->uint16();
        $rowLast = $this->uint16();
        $colFirst = $this->uint16();
        $colLast = $this->uint16();

        $from = $this->cellRef($rowFirst, $colFirst, $relative);
        $to = $this->cellRef($rowLast, $colLast, $relative);
        $this->push($from . ':' . $to);
    }

    /**
     * Render one A1 reference from a BIFF8 row and column-with-flags field
     *
     * @param int $row
     * @param int $colField
     * @param bool $relative TRUE for tRefN/tAreaN, where offsets are cell-relative
     *
     * @return string
     */
    private function cellRef(int $row, int $colField, bool $relative): string
    {
        $rowRelative = (bool)($colField & 0x8000);
        $colRelative = (bool)($colField & 0x4000);
        $col = $colField & 0x00FF;

        if ($relative) {
            // offsets are signed and taken from the owning cell
            if ($rowRelative) {
                $row = $this->baseRow + self::signed($row, 16);
            }
            if ($colRelative) {
                $col = $this->baseCol + self::signed($col, 8);
            }
        }

        $rowPart = ($rowRelative ? '' : '$') . ($row + 1);
        $colPart = ($colRelative ? '' : '$') . Helper::colLetter($col + 1);

        return $colPart . $rowPart;
    }

    /**
     * @param bool $fixed TRUE for tFunc, where the arg count is implied by the function
     *
     * @return void
     */
    private function func(bool $fixed): void
    {
        if ($fixed) {
            $index = $this->uint16();
            $argCount = self::fixedArgCount($index);
            if ($argCount < 0) {
                $this->failed = true;

                return;
            }
        }
        else {
            $argCount = $this->byte();
            $index = $this->uint16();
        }
        // tFuncVar keeps the prompt/CE flags in the top bits of the index
        $index &= 0x7FFF;

        $name = BiffFunction::name($index);
        if ($name === null || count($this->stack) < $argCount) {
            $this->failed = true;

            return;
        }

        $args = [];
        for ($i = 0; $i < $argCount; $i++) {
            array_unshift($args, $this->pop());
        }
        $this->push($name . '(' . implode(',', $args) . ')');
    }

    /**
     * @return void
     */
    private function pushString(): void
    {
        [$value, $consumed] = BiffString::readShort($this->data, $this->pos);
        $this->pos += $consumed;
        $this->push('"' . str_replace('"', '""', $value) . '"');
    }

    /**
     * @return void
     */
    private function pushNumber(): void
    {
        $value = unpack('e', substr($this->data, $this->pos, 8))[1];
        $this->pos += 8;
        $this->push(self::numberText($value));
    }

    /**
     * @return void
     */
    private function pushBool(): void
    {
        $this->push($this->byte() ? 'TRUE' : 'FALSE');
    }

    /**
     * @return void
     */
    private function pushError(): void
    {
        $this->push(BiffRecord::ERROR_CODES[$this->byte()] ?? '#ERR');
    }

    /**
     * A named range is stored as an index into the NAME table, which the parser
     * does not carry, so the name itself is unavailable
     *
     * @return void
     */
    private function pushName(): void
    {
        $this->failed = true;
    }

    /**
     * @param string $operator
     *
     * @return void
     */
    private function binary(string $operator): void
    {
        if (count($this->stack) < 2) {
            $this->failed = true;

            return;
        }
        $right = $this->pop();
        $left = $this->pop();
        $this->push($left . $operator . $right);
    }

    /**
     * @param string $operator
     *
     * @return void
     */
    private function unaryPrefix(string $operator): void
    {
        if (!$this->stack) {
            $this->failed = true;

            return;
        }
        $this->push($operator . $this->pop());
    }

    /**
     * @param string $operator
     *
     * @return void
     */
    private function unarySuffix(string $operator): void
    {
        if (!$this->stack) {
            $this->failed = true;

            return;
        }
        $this->push($this->pop() . $operator);
    }

    /**
     * tParen restores the parentheses the user typed
     *
     * @return void
     */
    private function parenthesis(): void
    {
        if (!$this->stack) {
            $this->failed = true;

            return;
        }
        $this->push('(' . $this->pop() . ')');
    }

    /**
     * @param string $text
     *
     * @return void
     */
    private function push(string $text): void
    {
        $this->stack[] = $text;
    }

    /**
     * @return string
     */
    private function pop(): string
    {
        return (string)array_pop($this->stack);
    }

    /**
     * @return int
     */
    private function byte(): int
    {
        return ord($this->data[$this->pos++] ?? "\0");
    }

    /**
     * @return int
     */
    private function uint16(): int
    {
        $value = unpack('v', substr($this->data, $this->pos, 2))[1];
        $this->pos += 2;

        return $value;
    }

    /**
     * Two's complement interpretation over $bits bits
     *
     * @param int $value
     * @param int $bits
     *
     * @return int
     */
    private static function signed(int $value, int $bits): int
    {
        $mask = 1 << ($bits - 1);
        $value &= (1 << $bits) - 1;

        return ($value & $mask) ? $value - (1 << $bits) : $value;
    }

    /**
     * Argument count of the fixed-arity built-in functions this parser renders;
     * -1 marks one it will not attempt
     *
     * @param int $index
     *
     * @return int
     */
    private static function fixedArgCount(int $index): int
    {
        static $counts = [
            2 => 1, 3 => 1, 10 => 0, 13 => 1, 14 => 1, 15 => 1, 16 => 1, 17 => 1,
            18 => 1, 19 => 0, 20 => 1, 21 => 1, 22 => 1, 23 => 1, 24 => 1, 25 => 1,
            26 => 1, 27 => 2, 30 => 2, 31 => 3, 32 => 1, 33 => 1, 34 => 0, 35 => 0,
            38 => 1, 39 => 2, 40 => 3, 41 => 3, 42 => 3, 43 => 3, 44 => 3, 45 => 3,
            47 => 3, 63 => 0, 65 => 3, 66 => 3, 67 => 1, 68 => 1, 69 => 1, 70 => 1,
            71 => 1, 72 => 1, 73 => 1, 74 => 0, 75 => 1, 76 => 1, 77 => 1, 97 => 2,
            98 => 1, 99 => 1, 105 => 1, 111 => 1, 112 => 1, 113 => 1, 114 => 1,
            117 => 2, 118 => 1, 119 => 4, 120 => 3, 121 => 1, 124 => 2, 126 => 1,
            127 => 1, 128 => 1, 129 => 1, 130 => 1, 131 => 1, 142 => 3, 143 => 4,
            144 => 4, 189 => 3, 190 => 1, 197 => 2, 198 => 1, 199 => 3, 212 => 2,
            213 => 2, 220 => 2, 221 => 0, 229 => 1, 230 => 1, 231 => 1, 232 => 1,
            233 => 1, 234 => 1, 235 => 3, 337 => 2, 342 => 1, 343 => 1,
        ];

        return $counts[$index] ?? -1;
    }

    /**
     * @param float $value
     *
     * @return string
     */
    private static function numberText(float $value): string
    {
        if (floor($value) === $value && abs($value) < 1e15) {
            return (string)(int)$value;
        }

        return rtrim(rtrim(sprintf('%.15G', $value), '0'), '.');
    }
}
