<?php

namespace avadim\FastExcelReader\Xls;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\AbstractSheet;

/**
 * XLS (BIFF8) worksheet reader
 *
 * Supplies AbstractSheet with raw rows. Cells arrive in ascending row order
 * inside the sheet substream, so a row is complete as soon as a record for a
 * later row appears - which is what makes a single forward pass enough and
 * keeps memory flat.
 */
class XlsSheet extends AbstractSheet
{
    private XlsBook $book;

    /** Absolute offset of this sheet's BOF inside the workbook stream */
    private int $bofOffset;

    /** Master token arrays of shared formulas, keyed "row:col" of the anchor */
    private ?array $sharedFormulaTokens = null;

    /**
     * @param string $name
     * @param string $sheetId
     * @param int $bofOffset
     * @param XlsBook $book
     */
    public function __construct(string $name, string $sheetId, int $bofOffset, XlsBook $book)
    {
        $this->excel = $book;
        $this->book = $book;
        $this->name = $name;
        $this->sheetId = $sheetId;
        $this->bofOffset = $bofOffset;

        $this->initReadArea();
    }

    /**
     * Offset of the sheet substream in the workbook stream
     *
     * @return string
     */
    public function path(): string
    {
        return (string)$this->bofOffset;
    }

    /**
     * A record reader positioned at the start of this sheet
     *
     * @return BiffReader
     */
    private function reader(): BiffReader
    {
        $biff = $this->book->newBiffReader();
        $biff->seek($this->bofOffset);
        $biff->nextRecord(); // the BOF of the substream

        return $biff;
    }

    /**
     * Read DIMENSIONS, which sits near the start of the substream
     *
     * @return void
     */
    protected function _readHeader()
    {
        if (isset($this->dimension['range'])) {
            return;
        }
        $this->dimension = ['range' => ''];

        foreach ($this->reader()->records() as $record) {
            if ($record['type'] === BiffRecord::DIMENSIONS) {
                $bounds = unpack('VrowFirst/VrowLast/vcolFirst/vcolLast', substr($record['data'], 0, 12));
                // DIMENSIONS is zero based and its last row/column are exclusive
                if ($bounds['rowLast'] > $bounds['rowFirst'] && $bounds['colLast'] > $bounds['colFirst']) {
                    $range = Helper::colLetter($bounds['colFirst'] + 1) . ($bounds['rowFirst'] + 1)
                        . ':' . Helper::colLetter($bounds['colLast']) . $bounds['rowLast'];
                    $this->dimension = Helper::rangeArray($range);
                    $this->dimension['range'] = $range;
                }
                break;
            }
            if ($record['type'] === BiffRecord::EOF) {
                break;
            }
        }
    }

    /**
     * Merge definitions live in MERGEDCELLS records inside the sheet substream
     *
     * @return array|null
     */
    public function getMergedCells(): ?array
    {
        if ($this->mergedCells !== null) {
            return $this->mergedCells;
        }
        $this->mergedCells = [];

        foreach ($this->reader()->records() as $record) {
            if ($record['type'] === BiffRecord::EOF) {
                break;
            }
            if ($record['type'] !== BiffRecord::MERGEDCELLS) {
                continue;
            }
            $count = unpack('v', substr($record['data'], 0, 2))[1];
            for ($i = 0; $i < $count; $i++) {
                $entry = unpack('vrowFirst/vrowLast/vcolFirst/vcolLast', substr($record['data'], 2 + $i * 8, 8));
                $range = Helper::colLetter($entry['colFirst'] + 1) . ($entry['rowFirst'] + 1)
                    . ':' . Helper::colLetter($entry['colLast'] + 1) . ($entry['rowLast'] + 1);
                $this->mergedCells[Helper::colLetter($entry['colFirst'] + 1) . ($entry['rowFirst'] + 1)] = $range;
            }
        }

        return $this->mergedCells;
    }

    /**
     * @param array $readArea
     * @param int $rowLimit
     * @param bool $rowMode
     *
     * @return \Generator|null
     */
    protected function rawRows(array $readArea, int $rowLimit = 0, bool $rowMode = false): ?\Generator
    {
        $biff = $this->reader();

        $currentRow = null;
        $cells = [];
        $rowAttributes = [];
        $rowCnt = -1;
        $pendingFormula = null;

        foreach ($biff->records() as $record) {
            $type = $record['type'];

            if ($type === BiffRecord::EOF) {
                break;
            }

            // the cached string result of the formula just seen
            if ($type === BiffRecord::STRING && $pendingFormula !== null) {
                [$value] = BiffString::readLong($record['data'], 0);
                $cells[Helper::colLetter($pendingFormula['col'])] = $this->makeCell($value, $pendingFormula['style'], 'string', $pendingFormula['formula']);
                $pendingFormula = null;

                continue;
            }
            if ($pendingFormula !== null) {
                // a string-result formula with no cached STRING (an uncalculated
                // file): keep the cell and its text rather than dropping it
                if ($pendingFormula['col'] >= $readArea['col_min'] && $pendingFormula['col'] <= $readArea['col_max']) {
                    $cells[Helper::colLetter($pendingFormula['col'])] = $this->makeCell('', $pendingFormula['style'], 'string', $pendingFormula['formula']);
                }
                $pendingFormula = null;
            }

            if ($type === BiffRecord::ROW) {
                if ($rowMode) {
                    $row = unpack('v', substr($record['data'], 0, 2))[1] + 1;
                    $rowAttributes[$row] = $this->rowAttributes($record['data']);
                }

                continue;
            }

            $parsed = $this->parseCellRecord($record, $pendingFormula);
            if ($parsed === null) {
                continue;
            }
            [$rowNum, $rowCells] = $parsed;

            if ($rowNum < $readArea['row_min']) {
                continue;
            }
            if ($rowNum > $readArea['row_max']) {
                break;
            }

            if ($currentRow !== null && $rowNum !== $currentRow) {
                $rowCnt++;
                if ($rowLimit > 0 && $rowCnt >= $rowLimit) {
                    return null;
                }
                yield $currentRow => [
                    'cells' => $cells,
                    'attrs' => $rowAttributes[$currentRow] ?? [],
                ];
                $cells = [];
            }
            $currentRow = $rowNum;

            foreach ($rowCells as $colNum => $cellData) {
                if ($colNum >= $readArea['col_min'] && $colNum <= $readArea['col_max']) {
                    $cells[Helper::colLetter($colNum)] = $cellData;
                }
            }
        }

        if ($currentRow !== null) {
            $rowCnt++;
            if ($rowLimit <= 0 || $rowCnt < $rowLimit) {
                yield $currentRow => [
                    'cells' => $cells,
                    'attrs' => $rowAttributes[$currentRow] ?? [],
                ];
            }
        }

        return null;
    }

    /**
     * Decode one cell-bearing record
     *
     * @param array $record
     * @param array|null $pendingFormula Set when a FORMULA expects a STRING record next
     *
     * @return array|null [rowNumber, [oneBasedColumn => cellData]] or NULL if the record holds no cells
     */
    private function parseCellRecord(array $record, ?array &$pendingFormula): ?array
    {
        $data = $record['data'];
        if (strlen($data) < 6) {
            return null;
        }
        $row = unpack('v', substr($data, 0, 2))[1] + 1;
        $col = unpack('v', substr($data, 2, 2))[1] + 1;
        $style = unpack('v', substr($data, 4, 2))[1];

        switch ($record['type']) {
            case BiffRecord::LABELSST:
                $index = unpack('V', substr($data, 6, 4))[1];

                return [$row, [$col => $this->makeCell($this->book->sharedString($index), $style, 'string')]];

            case BiffRecord::LABEL:
            case BiffRecord::RSTRING:
                [$value] = BiffString::readLong($data, 6);

                return [$row, [$col => $this->makeCell($value, $style, 'string')]];

            case BiffRecord::NUMBER:
                $value = unpack('e', substr($data, 6, 8))[1];

                return [$row, [$col => $this->makeCell($value, $style, 'number')]];

            case BiffRecord::RK:
                $value = self::decodeRk(unpack('V', substr($data, 6, 4))[1]);

                return [$row, [$col => $this->makeCell($value, $style, 'number')]];

            case BiffRecord::MULRK:
                $cells = [];
                $count = intdiv(strlen($data) - 6, 6);
                for ($i = 0; $i < $count; $i++) {
                    $entryStyle = unpack('v', substr($data, 4 + $i * 6, 2))[1];
                    $value = self::decodeRk(unpack('V', substr($data, 6 + $i * 6, 4))[1]);
                    $cells[$col + $i] = $this->makeCell($value, $entryStyle, 'number');
                }

                return [$row, $cells];

            case BiffRecord::BLANK:
                return [$row, [$col => $this->makeCell(null, $style, 'string')]];

            case BiffRecord::MULBLANK:
                $cells = [];
                $count = intdiv(strlen($data) - 6, 2);
                for ($i = 0; $i < $count; $i++) {
                    $entryStyle = unpack('v', substr($data, 4 + $i * 2, 2))[1];
                    $cells[$col + $i] = $this->makeCell(null, $entryStyle, 'string');
                }

                return [$row, $cells];

            case BiffRecord::BOOLERR:
                $raw = ord($data[6]);
                $isError = ord($data[7]) === 1;
                if ($isError) {
                    return [$row, [$col => $this->makeCell(BiffRecord::ERROR_CODES[$raw] ?? '#ERR', $style, 'error')]];
                }

                return [$row, [$col => $this->makeCell((bool)$raw, $style, 'bool')]];

            case BiffRecord::FORMULA:
                return [$row, $this->parseFormula($data, $row, $col, $style, $pendingFormula)];
        }

        return null;
    }

    /**
     * The cached result of a formula
     *
     * A result of 0xFFFF in the top two bytes marks a non-numeric result, and
     * the first byte then says which kind. A string result is not stored here
     * at all: it follows in a separate STRING record.
     *
     * @param string $data
     * @param int $row
     * @param int $col
     * @param int $style
     * @param array|null $pendingFormula
     *
     * @return array
     */
    private function parseFormula(string $data, int $row, int $col, int $style, ?array &$pendingFormula): array
    {
        $result = substr($data, 6, 8);
        // parseCellRecord() reports row and col one-based; the token decompiler
        // works in the zero-based coordinates the file itself uses
        $formula = $this->formulaText($data, $row - 1, $col - 1);

        if (substr($result, 6, 2) === "\xFF\xFF") {
            $kind = ord($result[0]);
            switch ($kind) {
                case 0: // string, carried by the next STRING record
                    $pendingFormula = ['col' => $col, 'style' => $style, 'formula' => $formula];

                    return [];

                case 1:
                    return [$col => $this->makeCell((bool)ord($result[2]), $style, 'bool', $formula)];

                case 2:
                    return [$col => $this->makeCell(BiffRecord::ERROR_CODES[ord($result[2])] ?? '#ERR', $style, 'error', $formula)];

                default: // empty string result
                    return [$col => $this->makeCell('', $style, 'string', $formula)];
            }
        }

        return [$col => $this->makeCell(unpack('e', $result)[1], $style, 'number', $formula)];
    }

    /**
     * Reconstruct the A1 text of a formula, or null if it cannot be rendered
     *
     * The token array either holds the formula directly, or is a single tExp
     * token pointing at the anchor of a shared formula whose master tokens live
     * in a SHRFMLA record. Either way the tokens are decompiled relative to this
     * cell, so B3 of a shared "=A+1" comes out as "=A3+1".
     *
     * @param string $data
     * @param int $row
     * @param int $col
     *
     * @return string|null
     */
    private function formulaText(string $data, int $row, int $col): ?string
    {
        $cce = unpack('v', substr($data, 20, 2))[1];
        $tokens = substr($data, 22, $cce);
        if ($tokens === '') {
            return null;
        }

        // tExp: the real tokens are shared, keyed by the anchor it names
        if (ord($tokens[0]) === 0x01) {
            $anchorRow = unpack('v', substr($tokens, 1, 2))[1];
            $anchorCol = unpack('v', substr($tokens, 3, 2))[1];
            $tokens = $this->sharedFormula($anchorRow, $anchorCol);
            if ($tokens === null) {
                return null;
            }
        }

        return (new FormulaParser($row, $col))->parse($tokens);
    }

    /**
     * Master token array of the shared formula anchored at (row, col)
     *
     * @param int $row
     * @param int $col
     *
     * @return string|null
     */
    private function sharedFormula(int $row, int $col): ?string
    {
        if ($this->sharedFormulaTokens === null) {
            $this->loadSharedFormulas();
        }

        return $this->sharedFormulaTokens[$row . ':' . $col] ?? null;
    }

    /**
     * Collect every SHRFMLA master token array in the sheet
     *
     * A SHRFMLA record follows the first cell of its group, so a forward scan
     * gathers them all. The scan is cheap and cached, and only runs when a
     * formula's text is actually requested.
     *
     * @return void
     */
    private function loadSharedFormulas(): void
    {
        $this->sharedFormulaTokens = [];

        foreach ($this->reader()->records() as $record) {
            if ($record['type'] === BiffRecord::EOF) {
                break;
            }
            if ($record['type'] !== BiffRecord::SHRFMLA) {
                continue;
            }
            $data = $record['data'];
            $rowFirst = unpack('v', substr($data, 0, 2))[1];
            $colFirst = ord($data[4]);
            $cce = unpack('v', substr($data, 8, 2))[1];
            $this->sharedFormulaTokens[$rowFirst . ':' . $colFirst] = substr($data, 10, $cce);
        }
    }

    /**
     * Build the cell descriptor shared by every reader in this library
     *
     * Number formats decide whether a numeric cell is really a date, exactly as
     * in the XLSX reader: the style index is looked up and its format category
     * drives the conversion.
     *
     * @param mixed $value
     * @param int $styleIdx
     * @param string $dataType
     * @param string|null $formula
     *
     * @return array
     */
    private function makeCell($value, int $styleIdx, string $dataType, ?string $formula = null): array
    {
        $originalValue = $value;

        if ($dataType === 'number' && $styleIdx > 0) {
            // the number format decides the type, exactly as in the XLSX reader:
            // a serial number under a date format is a date, and a number under
            // the text format "@" is a string
            $style = $this->excel->styleByIdx($styleIdx);
            $formatType = $style['formatType'] ?? null;

            if ($formatType === 'd' || $formatType === 'date') {
                $dataType = 'date';
                $formatter = $this->excel->getDateFormatter();
                if ($formatter === null) {
                    // dates are left as the original serial value
                }
                elseif ($formatter === false) {
                    $value = $this->excel->timestamp($value);
                }
                elseif ($timestamp = $this->excel->timestamp($value)) {
                    $value = $this->excel->formatDate($timestamp, null, $styleIdx);
                }
            }
            elseif ($formatType === 'string') {
                $dataType = 'string';
                $value = self::numberToString($value);
                $originalValue = $value;
            }
        }

        if ($dataType === 'number' && is_float($value) && floor($value) === $value && abs($value) < PHP_INT_MAX) {
            // whole numbers are reported as integers, as the XLSX reader does
            $value = (int)$value;
            $originalValue = $value;
        }

        return [
            'v' => $value,
            's' => $styleIdx,
            'f' => $formula,
            't' => $dataType,
            'o' => $originalValue,
        ];
    }

    /**
     * Attributes of a ROW record, named as in the XLSX reader where they overlap
     *
     * @param string $data
     *
     * @return array
     */
    private function rowAttributes(string $data): array
    {
        $fields = unpack('vrow/vcolFirst/vcolLast/vheight', substr($data, 0, 8));
        // grbit is a 32 bit field: the style index sits in bits 16..27
        $flags = unpack('V', substr($data, 12, 4))[1];

        $attributes = [
            'r' => (string)($fields['row'] + 1),
            'ht' => (string)(($fields['height'] & 0x7FFF) / 20),
        ];
        if ($flags & 0x0020) {
            $attributes['hidden'] = '1';
        }
        if ($flags & 0x0080) {
            $attributes['customFormat'] = '1';
            $attributes['s'] = (string)(($flags >> 16) & 0x0FFF);
        }

        return $attributes;
    }

    /**
     * Render a number the way it was written, so that a text-formatted 12345
     * reads back as "12345" and not as "12345.0"
     *
     * @param float|int $value
     *
     * @return string
     */
    private static function numberToString($value): string
    {
        if (is_float($value) && floor($value) === $value && abs($value) < PHP_INT_MAX) {
            return (string)(int)$value;
        }

        return (string)$value;
    }

    /**
     * Decode an RK number
     *
     * Two flag bits are stolen from the low end: one says the remaining 30 bits
     * are a signed integer rather than the high half of an IEEE double, the
     * other that the result must be divided by 100.
     *
     * @param int $rk
     *
     * @return float|int
     */
    private static function decodeRk(int $rk)
    {
        if ($rk & 0x02) {
            // arithmetic shift keeps the sign of the 30 bit integer
            $value = unpack('l', pack('V', $rk))[1] >> 2;
        }
        else {
            $value = unpack('e', "\x00\x00\x00\x00" . pack('V', $rk & 0xFFFFFFFC))[1];
        }

        if ($rk & 0x01) {
            $value /= 100;
        }

        return $value;
    }
}
