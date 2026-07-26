<?php

namespace avadim\FastExcelReader\Csv;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\AbstractSheet;

/**
 * CSV worksheet reader
 *
 * A CSV file is presented as a single worksheet. This class supplies
 * AbstractSheet with raw rows, tokenised by the CsvReader engine, so the whole
 * public reading API - key modes, read areas, result-mode flags, withHeader(),
 * the generator lifecycle - is inherited unchanged and behaves exactly as it
 * does for XLSX and XLS. Values are always strings: CSV carries no cell formats,
 * so there is no number/date typing and no styles.
 */
class CsvSheet extends AbstractSheet
{
    protected CsvReader $reader;

    protected string $file;

    /** TRUE once the actual extent has been scanned and cached into $dimension */
    protected bool $extentScanned = false;

    /**
     * @param string $name
     * @param string $sheetId
     * @param CsvReader $reader
     * @param string $file
     * @param CsvBook $book
     */
    public function __construct(string $name, string $sheetId, CsvReader $reader, string $file, CsvBook $book)
    {
        $this->excel = $book;
        $this->name = $name;
        $this->sheetId = $sheetId;
        $this->reader = $reader;
        $this->file = $file;

        $this->initReadArea();
    }

    /**
     * The CSV file itself is the sheet
     *
     * @return string
     */
    public function path(): string
    {
        return $this->file;
    }

    /**
     * The engine that tokenises the CSV file
     *
     * @return CsvReader
     */
    public function getReader(): CsvReader
    {
        return $this->reader;
    }

    /**
     * CSV stores no dimension, so a cheap sentinel is set here and the streaming
     * path stays scan free. The real extent is computed lazily, only when a
     * caller actually asks for it (see computeExtent()).
     *
     * @return void
     */
    protected function _readHeader()
    {
        if (isset($this->dimension['range'])) {
            return;
        }
        $this->dimension = ['range' => ''];
        $this->active = true;
    }

    /**
     * CSV has no merged cells
     *
     * @return array|null
     */
    public function getMergedCells(): ?array
    {
        return [];
    }

    /**
     * One streaming pass over the file, keyed by physical line number
     *
     * @param array $readArea
     * @param int $rowLimit
     * @param bool $rowMode
     *
     * @return \Generator|null
     */
    protected function rawRows(array $readArea, int $rowLimit = 0, bool $rowMode = false): ?\Generator
    {
        $this->reader->rewind();

        $yielded = 0;
        while (($fields = $this->reader->nextRawRow()) !== false) {
            $rowNum = $this->reader->getLineNo();
            if ($rowNum < $readArea['row_min']) {
                continue;
            }
            if ($rowNum > $readArea['row_max']) {
                break;
            }
            if ($rowLimit > 0 && $yielded >= $rowLimit) {
                break;
            }

            $cells = [];
            $colNum = 0;
            foreach ($fields as $field) {
                $colNum++;
                if ($colNum < $readArea['col_min']) {
                    continue;
                }
                if ($colNum > $readArea['col_max']) {
                    break;
                }
                $cells[Helper::colLetter($colNum)] = [
                    'v' => $field,
                    's' => null,
                    'f' => null,
                    't' => 's',
                    'o' => $field,
                ];
            }

            $yielded++;
            yield $rowNum => [
                'cells' => $cells,
                'attrs' => $rowMode ? ['r' => (string)$rowNum] : [],
            ];
        }

        return null;
    }

    /**
     * Scan the whole file once to establish the real extent, and cache it
     *
     * Kept out of the streaming path on purpose: nextRow() only touches the
     * cheap sentinel from _readHeader(), so reading a huge CSV never forces this
     * counting pass. It runs only when a caller asks for the extent explicitly
     * (dimension()/countRows()/...), and memory stays flat because nothing is
     * materialised - rows are only counted.
     *
     * @return void
     */
    protected function computeExtent(): void
    {
        if ($this->extentScanned) {
            return;
        }
        $this->extentScanned = true;

        // The shared engine is reused (rewound), not a second reader, so the scan
        // sees exactly the same options - delimiter, encoding, skip-empty-lines,
        // comment prefix - as rawRows(). A fresh CsvReader built from getOptions()
        // would drop the flags getOptions() omits and could count rows differently.
        $this->reader->rewind();
        $minRow = $maxRow = $maxCol = 0;
        while (($fields = $this->reader->nextRawRow()) !== false) {
            $rowNum = $this->reader->getLineNo();
            if ($minRow === 0) {
                $minRow = $rowNum;
            }
            $maxRow = $rowNum;
            $count = count($fields);
            if ($count > $maxCol) {
                $maxCol = $count;
            }
        }

        if ($maxRow === 0) {
            $this->dimension = ['range' => ''];

            return;
        }
        $range = 'A' . $minRow . ':' . Helper::colLetter($maxCol) . $maxRow;
        $this->dimension = Helper::rangeArray($range);
        $this->dimension['range'] = $range;
    }

    /**
     * Actual data range of the CSV, computed by scanning the file
     *
     * dimension() returns this too once it has been computed; unlike XLSX there
     * is no declared range to return before the first scan.
     *
     * @return string
     */
    public function actualDimension(): string
    {
        $this->computeExtent();

        return (string)$this->dimension['range'];
    }

    /**
     * @return array
     */
    public function dimensionArray(): array
    {
        $this->computeExtent();

        return $this->dimension;
    }

    /**
     * @param string|null $range
     *
     * @return int
     */
    public function countRows(?string $range = null): int
    {
        if ($range === null) {
            $this->computeExtent();
        }

        return parent::countRows($range);
    }

    /**
     * @param string|null $range
     *
     * @return int
     */
    public function countColumns(?string $range = null): int
    {
        if ($range === null) {
            $this->computeExtent();
        }

        return parent::countColumns($range);
    }

    /**
     * @param string|null $range
     *
     * @return int
     */
    public function minRow(?string $range = null): int
    {
        if ($range === null) {
            $this->computeExtent();
        }

        return parent::minRow($range);
    }

    /**
     * @param string|null $range
     *
     * @return int
     */
    public function maxRow(?string $range = null): int
    {
        if ($range === null) {
            $this->computeExtent();
        }

        return parent::maxRow($range);
    }

    /**
     * @param string|null $range
     *
     * @return string
     */
    public function minColumn(?string $range = null): string
    {
        if ($range === null) {
            $this->computeExtent();
        }

        return parent::minColumn($range);
    }

    /**
     * @param string|null $range
     *
     * @return string
     */
    public function maxColumn(?string $range = null): string
    {
        if ($range === null) {
            $this->computeExtent();
        }

        return parent::maxColumn($range);
    }
}
