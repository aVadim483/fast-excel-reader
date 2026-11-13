<?php

namespace avadim\FastExcelReader\Ods;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\AbstractSheet;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Reader;

/** Reads one table from the shared content.xml without expanding empty repetitions. */
class OdsSheet extends AbstractSheet
{
    public function __construct(string $name, string $sheetId, OdsBook $book)
    {
        $this->excel = $book;
        $this->name = $name;
        $this->sheetId = $sheetId;
        $this->initReadArea();
    }

    public function path(): string
    {
        return 'content.xml#' . $this->sheetId;
    }

    /** Keep streaming reads scan-free; actual dimensions are computed only on request. */
    protected function _readHeader()
    {
        $this->dimension = ['range' => ''];
        $this->active = $this->sheetId === '1';
    }

    private function repeat(Reader $reader, string $name): int
    {
        $value = $reader->getAttributeNs($name, OdsBook::TABLE_NS);
        if ($value === null) {
            return 1;
        }
        if (!preg_match('/^[1-9][0-9]*$/D', $value) || strlen($value) > 9 || (int)$value > 1048576) {
            throw new Exception('Invalid or unsupported ODS repetition/span: ' . $value);
        }
        return (int)$value;
    }

    /** Yield compact row runs; a repeated row is represented once. */
    private function rowRuns(int $columnMin, int $columnMax, int $rowMax = 1048576, bool $metadata = false, int $rowMin = 1): \Generator
    {
        $reader = $this->excel->openContent();
        $table = 0;
        $tableDepth = null;
        $row = 1;
        $spreadsheetDepth = -1;
        try {
            while ($reader->read()) {
                OdsBook::checkNode($reader);
                if ($reader->nodeType === \XMLReader::ELEMENT && $reader->namespaceURI === OdsBook::OFFICE_NS && $reader->localName === 'spreadsheet') {
                    $spreadsheetDepth = $reader->depth;
                }
                if ($reader->namespaceURI !== OdsBook::TABLE_NS) {
                    continue;
                }
                if ($reader->localName === 'table' && $reader->nodeType === \XMLReader::ELEMENT && $reader->depth === $spreadsheetDepth + 1) {
                    $table++;
                    if ($table === (int)$this->sheetId) {
                        $tableDepth = $reader->depth;
                        if ($reader->isEmptyElement) {
                            return;
                        }
                    }
                }
                if ($tableDepth === null) {
                    continue;
                }
                if ($reader->nodeType === \XMLReader::END_ELEMENT && $reader->localName === 'table' && $reader->depth === $tableDepth) {
                    return;
                }
                if ($reader->nodeType !== \XMLReader::ELEMENT || $reader->localName !== 'table-row') {
                    continue;
                }
                if ($row > $rowMax) {
                    return;
                }
                $rows = $this->repeat($reader, 'number-rows-repeated');
                if ($row + $rows - 1 > 1048576) {
                    throw new Exception('ODS row limit exceeded');
                }
                $depth = $reader->depth;
                $cells = $spans = [];
                $col = 1;
                $maxUsed = 0;
                $minUsed = PHP_INT_MAX;
                if (!$reader->isEmptyElement) {
                    while ($reader->read()) {
                        OdsBook::checkNode($reader);
                        if ($reader->nodeType === \XMLReader::END_ELEMENT && $reader->depth === $depth) {
                            break;
                        }
                        if ($reader->nodeType !== \XMLReader::ELEMENT || $reader->namespaceURI !== OdsBook::TABLE_NS || !in_array($reader->localName, ['table-cell', 'covered-table-cell'], true)) {
                            continue;
                        }
                        $columns = $this->repeat($reader, 'number-columns-repeated');
                        if ($col + $columns - 1 > 16384) {
                            throw new Exception('ODS column limit exceeded');
                        }
                        $width = $this->repeat($reader, 'number-columns-spanned');
                        $height = $this->repeat($reader, 'number-rows-spanned');
                        if ($col + $columns + $width - 2 > 16384 || $row + $rows + $height - 2 > 1048576) {
                            throw new Exception('ODS merged range exceeds supported dimensions');
                        }
                        $covered = $reader->localName === 'covered-table-cell';
                        $intersects = $col <= $columnMax && $col + $columns - 1 >= $columnMin;
                        $value = null;
                        if (!$covered && (($intersects && $row + $rows - 1 >= $rowMin) || $metadata)) {
                            $node = $reader->expand();
                            if (!$node) {
                                throw new Exception('Cannot read ODS cell');
                            }
                            $value = $metadata ? $this->cellPresent($node) : $this->cell($node);
                        }
                        if ($value !== null && $value !== false) {
                            $minUsed = min($minUsed, $col);
                            $maxUsed = max($maxUsed, $col + $columns - 1, $col + $width - 1);
                            if (!$metadata) {
                                for ($c = max($col, $columnMin); $c <= min($col + $columns - 1, $columnMax); $c++) {
                                    $cells[Helper::colLetter($c)] = $value;
                                }
                            }
                        }
                        if ($metadata && ($width > 1 || $height > 1)) {
                            $minUsed = min($minUsed, $col);
                            $spans[] = [$col, $columns, $width, $height];
                            $maxUsed = max($maxUsed, $col + $columns + $width - 2);
                        }
                        // Skip cell children, including embedded nested tables.
                        if (!$reader->isEmptyElement) {
                            $cellDepth = $reader->depth;
                            while ($reader->read()) {
                                OdsBook::checkNode($reader);
                                if ($reader->nodeType === \XMLReader::END_ELEMENT && $reader->depth === $cellDepth) {
                                    break;
                                }
                            }
                        }
                        $col += $columns;
                    }
                }
                yield ['row' => $row, 'repeat' => $rows, 'cells' => $cells, 'minCol' => $minUsed, 'maxCol' => $maxUsed, 'spans' => $spans];
                $row += $rows;
            }
            throw new Exception('Incomplete ODS worksheet XML');
        }
        finally {
            $reader->close();
            $this->excel->releaseReader($reader);
        }
    }

    private function cellPresent(\DOMElement $node): bool
    {
        if ($node->hasAttributeNS(OdsBook::OFFICE_NS, 'value-type') || $node->hasAttributeNS(OdsBook::TABLE_NS, 'formula')) {
            return true;
        }
        foreach ($node->childNodes as $child) {
            if ($child instanceof \DOMElement && $child->namespaceURI === OdsBook::TEXT_NS && $child->localName === 'p') {
                return true;
            }
        }
        return false;
    }

    /** Text paragraphs preserve explicit ODF spaces, tabs and line breaks. */
    private function text(\DOMNode $node): string
    {
        $text = '';
        foreach ($node->childNodes as $child) {
            if ($child->nodeType === XML_TEXT_NODE || $child->nodeType === XML_CDATA_SECTION_NODE) {
                $text .= $child->nodeValue;
            }
            elseif ($child instanceof \DOMElement) {
                if ($child->namespaceURI === OdsBook::TEXT_NS && $child->localName === 's') {
                    $count = $child->getAttributeNS(OdsBook::TEXT_NS, 'c') ?: '1';
                    if (!ctype_digit($count) || strlen($count) > 7 || (int)$count > 1000000) {
                        throw new Exception('Invalid ODS space repetition');
                    }
                    $text .= str_repeat(' ', (int)$count);
                }
                elseif ($child->namespaceURI === OdsBook::TEXT_NS && $child->localName === 'tab') {
                    $text .= "\t";
                }
                elseif ($child->namespaceURI === OdsBook::TEXT_NS && $child->localName === 'line-break') {
                    $text .= "\n";
                }
                else {
                    $text .= $this->text($child);
                }
            }
        }
        return $text;
    }

    private function cell(\DOMElement $node): ?array
    {
        if (!$this->cellPresent($node)) {
            return null;
        }
        $type = $node->getAttributeNS(OdsBook::OFFICE_NS, 'value-type');
        $original = null;
        switch ($type) {
            case 'float':
            case 'percentage':
            case 'currency':
                $original = $node->getAttributeNS(OdsBook::OFFICE_NS, 'value');
                if (!is_numeric($original)) {
                    throw new Exception('Invalid ODS numeric value: ' . $original);
                }
                $value = $original + 0;
                $type = 'n';
                break;
            case 'boolean':
                $original = $node->getAttributeNS(OdsBook::OFFICE_NS, 'boolean-value');
                if (!in_array($original, ['true', 'false', '1', '0'], true)) {
                    throw new Exception('Invalid ODS boolean value');
                }
                $value = $original === 'true' || $original === '1';
                $type = 'b';
                break;
            case 'date':
                $original = $node->getAttributeNS(OdsBook::OFFICE_NS, 'date-value');
                if ($original === '' || !preg_match('/^\d{4}-\d{2}-\d{2}(?:T.*)?$/D', $original)) {
                    throw new Exception('Invalid ODS date value');
                }
                $value = $this->excel->dateValue($original);
                $type = 'd';
                break;
            case 'time':
                // Preserve ISO duration text: negative or >24h values are not clock timestamps.
                $value = $original = $node->getAttributeNS(OdsBook::OFFICE_NS, 'time-value');
                if (!preg_match('/^-?P(?=\d|T\d)(?:\d+D)?(?:T(?:\d+H)?(?:\d+M)?(?:\d+(?:\.\d+)?S)?)?$/D', $value)) {
                    throw new Exception('Invalid ODS duration');
                }
                $type = 'duration';
                break;
            default:
                $paragraphs = [];
                foreach ($node->childNodes as $paragraph) {
                    if ($paragraph instanceof \DOMElement && $paragraph->namespaceURI === OdsBook::TEXT_NS && $paragraph->localName === 'p') {
                        $paragraphs[] = $this->text($paragraph);
                    }
                }
                $value = $original = $node->hasAttributeNS(OdsBook::OFFICE_NS, 'string-value')
                    ? $node->getAttributeNS(OdsBook::OFFICE_NS, 'string-value') : implode("\n", $paragraphs);
                if (!$paragraphs && $type === '' && !$node->hasAttributeNS(OdsBook::OFFICE_NS, 'string-value')) {
                    $value = $original = null;
                }
                $type = 's';
        }
        return ['v' => $value, 's' => null, 'f' => null, 't' => $type, 'o' => $original];
    }

    protected function rawRows(array $readArea, int $rowLimit = 0, bool $rowMode = false): ?\Generator
    {
        $yielded = 0;
        foreach ($this->rowRuns($readArea['col_min'], $readArea['col_max'], $readArea['row_max'], false, $readArea['row_min']) as $run) {
            if (!$run['cells']) {
                continue;
            }
            $start = max($run['row'], $readArea['row_min']);
            $end = min($run['row'] + $run['repeat'] - 1, $readArea['row_max']);
            for ($row = $start; $row <= $end; $row++) {
                $this->excel->assertOpen();
                if ($rowLimit > 0 && $yielded >= $rowLimit) {
                    return;
                }
                $yielded++;
                yield $row => ['cells' => $run['cells'], 'attrs' => $rowMode ? ['r' => (string)$row] : []];
            }
        }
    }

    private function computeMetadata(): void
    {
        $this->excel->assertOpen();
        if ($this->mergedCells !== null) {
            return;
        }
        $merges = [];
        $minRow = $minCol = PHP_INT_MAX;
        $maxRow = $maxCol = 0;
        foreach ($this->rowRuns(1, 16384, 1048576, true) as $run) {
            if ($run['maxCol']) {
                $minRow = min($minRow, $run['row']);
                $minCol = min($minCol, $run['minCol']);
                $maxRow = max($maxRow, $run['row'] + $run['repeat'] - 1);
                $maxCol = max($maxCol, $run['maxCol']);
            }
            foreach ($run['spans'] as [$col, $columns, $width, $height]) {
                for ($r = $run['row']; $r < $run['row'] + $run['repeat']; $r++) {
                    for ($c = $col; $c < $col + $columns; $c++) {
                        $range = Helper::colLetter($c) . $r . ':' . Helper::colLetter($c + $width - 1) . ($r + $height - 1);
                        $merges[$range] = $range;
                        $maxRow = max($maxRow, $r + $height - 1);
                    }
                }
            }
        }
        $this->mergedCells = $merges;
        $range = $maxRow ? Helper::colLetter($minCol) . $minRow . ':' . Helper::colLetter($maxCol) . $maxRow : '';
        $this->dimension = $range ? Helper::rangeArray($range) : [];
        $this->dimension['range'] = $range;
    }

    public function getMergedCells(): ?array
    {
        $this->computeMetadata();
        return $this->mergedCells;
    }

    public function actualDimension(): string
    {
        $this->computeMetadata();
        return $this->dimension['range'];
    }

    public function dimensionArray(): array
    {
        $this->computeMetadata();
        return $this->dimension;
    }

    public function countRows(?string $range = null): int
    {
        if ($range === null) {
            $this->computeMetadata();
        }
        return parent::countRows($range);
    }

    public function countColumns(?string $range = null): int
    {
        if ($range === null) {
            $this->computeMetadata();
        }
        return parent::countColumns($range);
    }

    public function minRow(?string $range = null): int
    {
        if ($range === null) {
            $this->computeMetadata();
        }
        return parent::minRow($range);
    }

    public function maxRow(?string $range = null): int
    {
        if ($range === null) {
            $this->computeMetadata();
        }
        return parent::maxRow($range);
    }

    public function minColumn(?string $range = null): string
    {
        if ($range === null) {
            $this->computeMetadata();
        }
        return parent::minColumn($range);
    }

    public function maxColumn(?string $range = null): string
    {
        if ($range === null) {
            $this->computeMetadata();
        }
        return parent::maxColumn($range);
    }
}
