<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceBookReader;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;

/**
 * Format independent part of a worksheet reader
 *
 * Everything here works in terms of rows, columns and read areas, and knows
 * nothing about how the sheet is stored. A concrete reader supplies the raw
 * rows and the sheet header; the whole public reading API - key modes, read
 * areas, result-mode flags, column renaming, the generator lifecycle - is
 * implemented once, here.
 *
 * @see Sheet for the XLSX implementation
 */
abstract class AbstractSheet implements InterfaceSheetReader
{
    public InterfaceBookReader $excel;

    protected string $sheetId;

    protected string $name;

    protected string $state = '';

    protected ?array $dimension = null;

    protected ?array $cols = null;

    protected ?bool $active = null;

    protected array $area = [];

    protected array $props = [];

    protected array $images = [];

    protected ?array $mergedCells = null;

    protected int $readRowNum = 0;

    /**
     * @var \Generator|null
     */
    protected ?\Generator $generator = null;

    protected int $countReadRows = 0;

    /**
     * Walk the sheet and yield one raw row at a time
     *
     * Yields $rowNum => ['cells' => [colLetter => cellData], 'attrs' => array]
     * where cellData is ['v' => value, 's' => styleIdx, 'f' => formula,
     * 't' => type, 'o' => original value]. Cells outside the column range of
     * $readArea must not be reported, so that their values are never parsed.
     *
     * @param array $readArea
     * @param int $rowLimit
     * @param bool $rowMode TRUE when row attributes are needed and empty rows must be reported
     *
     * @return \Generator|null
     */
    abstract protected function rawRows(array $readArea, int $rowLimit = 0, bool $rowMode = false): ?\Generator;

    /**
     * Populate $this->dimension, and optionally $this->active and $this->cols
     *
     * @return void
     */
    abstract protected function _readHeader();

    /**
     * Get merged cells. Returns an array [min_cell => range]
     *
     * @return array|null
     */
    abstract public function getMergedCells(): ?array;

    /**
     * Reset the read area to the whole sheet
     *
     * @return void
     */
    protected function initReadArea(): void
    {
        $this->area = [
            'row_min' => 1,
            'col_min' => 1,
            'row_max' => Helper::EXCEL_2007_MAX_ROW,
            'col_max' => Helper::EXCEL_2007_MAX_COL,
            'first_row_keys' => false,
            'col_keys' => [],
            'col_names' => [],
        ];
    }

    /**
     * Get sheet ID
     *
     * @return string
     */
    public function id(): string
    {
        return $this->sheetId;
    }

    /**
     * Get sheet name
     *
     * @return string
     */
    public function name(): string
    {
        return $this->name;
    }

    /**
     * Where the sheet lives inside its container
     *
     * An inner file path for XLSX, a stream offset for a format that stores
     * sheets as regions of one stream.
     *
     * @return string
     */
    abstract public function path(): string;

    /**
     * Case-insensitive name checking
     *
     * @param string $name
     *
     * @return bool
     */
    public function isName(string $name): bool
    {
        return strcasecmp($this->name, $name) === 0;
    }

    /**
     * Returns true if the sheet is active
     *
     * @return bool
     */
    public function isActive(): bool
    {
        if ($this->active === null) {
            $this->_readHeader();

            if ($this->active === null) {
                $this->active = false;
            }
        }

        return $this->active;
    }

    /**
     * Set sheet state (visible, hidden, veryHidden)
     *
     * @param string $state
     *
     * @return $this
     */
    public function setState(string $state): AbstractSheet
    {
        $this->state = $state;

        return $this;
    }

    /**
     * Get sheet state
     *
     * @return string
     */
    public function state(): string
    {
        return $this->state;
    }

    /**
     * Returns true if the sheet is visible
     *
     * @return bool
     */
    public function isVisible(): bool
    {
        return !$this->state || $this->state === 'visible';
    }

    /**
     * Returns true if the sheet is hidden
     *
     * @return bool
     */
    public function isHidden(): bool
    {
        return $this->state === 'hidden' || $this->state === 'veryHidden';
    }

    /**
     * Get sheet dimension range (e.g. "A1:C10")
     *
     * @return string|null
     */
    public function dimension(): ?string
    {
        if (!isset($this->dimension['range'])) {
            $this->_readHeader();
        }

        return $this->dimension['range'];
    }

    /**
     * Get sheet dimension as an array
     *
     * @return array
     */
    public function dimensionArray(): array
    {
        if (!isset($this->dimension['range'])) {
            $this->_readHeader();
        }

        return $this->dimension;
    }

    /**
     * Count rows by dimension value
     *
     * @param string|null $range
     *
     * @return int
     */
    public function countRows(?string $range = null): int
    {
        // A1:C3 || A1
        $areaRange = $range ?: $this->dimension();
        if ($areaRange && preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            return count($matches) === 6 ? ((int)$matches[5] - (int)$matches[2] + 1) : 1;
        }

        return 0;
    }

    /**
     * Count columns by dimension value
     *
     * @param string|null $range
     *
     * @return int
     */
    public function countColumns(?string $range = null): int
    {
        $areaRange = $range ?: $this->dimension();
        if ($areaRange && preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            return !empty($matches[4]) ? (Excel::colNum($matches[4]) - Excel::colNum($matches[1]) + 1) : 1;
        }

        return 0;
    }

    /**
     * Min row number from dimension value
     *
     * @param string|null $range
     *
     * @return int
     */
    public function minRow(?string $range = null): int
    {
        $areaRange = $range ?: $this->dimension();
        if ($areaRange && preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            return (int)$matches[2];
        }

        return 0;
    }

    /**
     * Max row number from dimension value
     *
     * @param string|null $range
     *
     * @return int
     */
    public function maxRow(?string $range = null): int
    {
        $areaRange = $range ?: $this->dimension();
        if ($areaRange && preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            return count($matches) === 6 ? (int)$matches[5] : (int)$matches[2];
        }

        return 0;
    }

    /**
     * Min column from dimension value
     *
     * @param string|null $range
     *
     * @return string
     */
    public function minColumn(?string $range = null): string
    {
        $areaRange = $range ?: $this->dimension();
        if ($areaRange && preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            return $matches[1] ?? '';
        }

        return '';
    }

    /**
     * Max column from dimension value
     *
     * @param string|null $range
     *
     * @return string
     */
    public function maxColumn(?string $range = null): string
    {
        $areaRange = $range ?: $this->dimension();
        if ($areaRange && preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            return $matches[4] ?? $this->minColumn($range);
        }

        return $this->minColumn($range);
    }

    /**
     * Count columns by dimension value, alias of countColumns()
     *
     * @param string|null $range
     *
     * @return int
     */
    public function countCols(?string $range = null): int
    {
        return $this->countColumns($range);
    }

    /**
     * Set date format
     *
     * @param $dateFormat
     *
     * @return $this
     */
    public function setDateFormat($dateFormat): AbstractSheet
    {
        $this->excel->setDateFormat($dateFormat);

        return $this;
    }

    protected static function _areaRange(string $areaRange): array
    {
        $area = [
            'col_keys' => [],
            'col_names' => [],
        ];
        if (preg_match('/^\$?([A-Za-z]+)\$?(\d+)(:\$?([A-Za-z]+)\$?(\d+))?$/', $areaRange, $matches)) {
            $area['col_min'] = Helper::colNumber($matches[1]);
            $area['row_min'] = (int)$matches[2];
            if (empty($matches[3])) {
                $area['col_max'] = Helper::EXCEL_2007_MAX_COL;
                $area['row_max'] = Helper::EXCEL_2007_MAX_ROW;
            }
            else {
                $area['col_max'] = Helper::colNumber($matches[4]);
                $area['row_max'] = (int)$matches[5];
                for ($col = $area['col_min']; $col <= $area['col_max']; $col++) {
                    $area['col_keys'][Helper::colLetter($col)] = null;
                }
            }
        }
        elseif (preg_match('/^([A-Za-z]+)(:([A-Za-z]+))?$/', $areaRange, $matches)) {
            $area['col_min'] = Helper::colNumber($matches[1]);
            if (empty($matches[2])) {
                $area['col_max'] = Helper::EXCEL_2007_MAX_COL;
            }
            else {
                $area['col_max'] = Helper::colNumber($matches[3]);
                for ($col = $area['col_min']; $col <= $area['col_max']; $col++) {
                    $area['col_keys'][Helper::colLetter($col)] = null;
                }
            }
            $area['row_min'] = 1;
            $area['row_max'] = Helper::EXCEL_2007_MAX_ROW;
        }
        if (isset($area['col_min'], $area['col_max']) && ($area['col_min'] < 0 || $area['col_max'] < 0)) {
            return [];
        }

        return $area;
    }

    /**
     * Set top left and right bottom of read area
     *
     * @param string $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return $this
     *
     * @example
     *  setReadArea('C3:AZ28'); // set top left and right bottom of read area
     *  setReadArea('C3'); // set top left only
     */
    public function setReadArea(string $areaRange, ?bool $firstRowKeys = false): AbstractSheet
    {
        if (preg_match('/^\w+$/', $areaRange)) {
            foreach ($this->excel->getDefinedNames() as $name => $range) {
                if ($name === $areaRange && strpos($range, $this->name . '!') === 0) {
                    [$sheetName, $definedRange] = explode('!', $range);
                    $areaRange = $definedRange;
                    break;
                }
            }
        }
        $area = self::_areaRange($areaRange);
        if ($area && isset($area['row_max'])) {
            $this->area = $area;
            $this->area['first_row_keys'] = $firstRowKeys;

            return $this;
        }
        throw new Exception('Wrong address or range "' . $areaRange . '"');
    }

    /**
     * Set top left of read area. Alias of setReadArea()
     *
     * @param string $topLeftCell
     * @param bool|null $firstRowKeys
     *
     * @return $this
     */
    public function from(string $topLeftCell, ?bool $firstRowKeys = false): AbstractSheet
    {
        if (strpos($topLeftCell, ':') !== false) {
            throw new Exception('Wrong address of top left cell "' . $topLeftCell . '"');
        }
        return $this->setReadArea($topLeftCell, $firstRowKeys);
    }

    /**
     * setReadArea('C:AZ') - set left and right columns of read area
     * setReadArea('C') - set left column only
     *
     * @param string $columnsRange
     * @param bool|null $firstRowKeys
     *
     * @return $this
     */
    public function setReadAreaColumns(string $columnsRange, ?bool $firstRowKeys = false): AbstractSheet
    {
        $area = self::_areaRange($columnsRange);
        if ($area) {
            $this->area = $area;
            $this->area['first_row_keys'] = $firstRowKeys;

            return $this;
        }
        throw new Exception('Wrong address or range "' . $columnsRange . '"');
    }

    /**
     * Enables header mode
     *
     * Treats the first row of the read area as a header row and returns subsequent rows
     * as associative arrays keyed by column names
     *
     * @return $this
     */
    public function withHeader(): AbstractSheet
    {
        $this->area['first_row_keys'] = true;

        return $this;
    }

    /**
     * Returns cell values as a two-dimensional array
     *      [1 => ['A' => _value_A1_], ['B' => _value_B1_]],
     *      [2 => ['A' => _value_A2_], ['B' => _value_B2_]]
     *
     *  readRows()
     *  readRows(true)
     *  readRows(false, Excel::KEYS_ZERO_BASED)
     *  readRows(Excel::KEYS_ZERO_BASED | Excel::KEYS_RELATIVE)
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readRows($columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null): array
    {
        $data = [];
        if (is_int($columnKeys) && !is_int($resultMode)) {
            $resultMode = $columnKeys;
            $columnKeys = [];
        }
        $this->readCallback(static function($row, $col, $val) use (&$columnKeys, &$data) {
            if (isset($columnKeys[$col])) {
                $data[$row][$columnKeys[$col]] = $val;
            }
            else {
                $data[$row][$col] = $val;
            }
        }, $columnKeys, $resultMode, $styleIdxInclude);

        if ($data && ($resultMode & Excel::KEYS_SWAP)) {
            $newData = [];
            $rowKeys = array_keys($data);
            $len = count($rowKeys);
            foreach (array_keys(reset($data)) as $colKey) {
                $rowValues = array_column($data, $colKey);
                if ($len - count($rowValues)) {
                    $rowValues = array_pad($rowValues, $len, null);
                }
                $newData[$colKey] = array_combine($rowKeys, $rowValues);
            }
            return $newData;
        }

        return $data;
    }

    /**
     * Read rows from a given area $areaRange
     *
     * @param string $areaRange
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readRowsFrom(string $areaRange, $columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null): array
    {
        return $this->setReadArea($areaRange)->readRows($columnKeys, $resultMode, $styleIdxInclude);
    }

    /**
     * Returns values, styles, and other info of cells as array
     *
     * [
     *      'v' => _value_,
     *      's' => _styles_,
     *      'f' => _formula_,
     *      't' => _type_,
     *      'o' => '_original_value_
     * ]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readRowsWithStyles($columnKeys = [], ?int $resultMode = null): array
    {
        $data = $this->readRows($columnKeys, $resultMode, true);

        foreach ($data as $row => $rowData) {
            foreach ($rowData as $col => $cellData) {
                if (isset($cellData['s'])) {
                    $data[$row][$col]['s'] = $this->excel->getCompleteStyleByIdx($cellData['s']);
                }
            }
        }

        return $data;
    }

    /**
     * Set read area and returns values, styles, and other info of cells as array
     *
     * @param string $areaRange
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readRowsWithStylesFrom(string $areaRange, $columnKeys = [], ?int $resultMode = null): array
    {
        return $this->setReadArea($areaRange)->readRowsWithStyles($columnKeys, $resultMode);
    }

    /**
     * Get number of the first row in the read area
     *
     * @return int
     */
    public function firstRow(): int
    {
        if (!isset($this->area['first_row'])) {
            $this->readFirstRow();
        }

        return $this->area['first_row'];
    }

    /**
     * Get letter of the first column in the read area
     *
     * @return string
     */
    public function firstCol(): string
    {
        if (!isset($this->area['first_col'])) {
            $this->readFirstRow();
        }

        return $this->area['first_col'];
    }

    /**
     * Returns values of cells of 1st row as array
     *
     * @param array|bool|int|null $columnKeys
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readFirstRow($columnKeys = [], ?bool $styleIdxInclude = null): array
    {
        $rowData = [];
        $rowNum = -1;
        $this->readCallback(static function($row, $col, $val) use (&$columnKeys, &$rowData, &$rowNum) {
            if ($rowNum === -1) {
                $rowNum = $row;
            }
            elseif ($rowNum !== $row) {
                return true;
            }
            if (isset($columnKeys[$col])) {
                $col = $rowData[$columnKeys[$col]];
            }
            $rowData[$col] = $val;

            return null;
        }, $columnKeys, null, $styleIdxInclude);

        return $rowData;
    }

    /**
     * Set read area and returns values of cells of 1st row as array
     *
     * @param string $areaRange
     * @param array|bool|int|null $columnKeys
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readFirstRowFrom(string $areaRange, $columnKeys = [], ?bool $styleIdxInclude = null): array
    {
        return $this->setReadArea($areaRange)->readFirstRow($columnKeys, $styleIdxInclude);
    }

    /**
     * Returns values and styles of cells of 1st row as array
     *
     * @param array|bool|int|null $columnKeys
     *
     * @return array
     */
    public function readFirstRowWithStyles($columnKeys = []): array
    {
        $rowData = $this->readFirstRow($columnKeys, true);
        foreach ($rowData as $col => $cellData) {
            if (isset($cellData['s'])) {
                $rowData[$col]['s'] = $this->excel->getCompleteStyleByIdx($cellData['s']);
            }
        }

        return $rowData;
    }

    /**
     * Set read area and returns values and styles of cells of 1st row as array
     *
     * @param string $areaRange
     * @param array|bool|int|null $columnKeys
     *
     * @return array
     */
    public function readFirstRowWithStylesFrom(string $areaRange, $columnKeys = []): array
    {
        return $this->setReadArea($areaRange)->readFirstRowWithStyles($columnKeys);
    }

    /**
     * Returns values and styles of cells of 1st row as array
     *
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readFirstRowCells(?bool $styleIdxInclude = null): array
    {
        $rowData = [];
        $rowNum = -1;
        $this->readCallback(static function($row, $col, $val) use (&$rowData, &$rowNum) {
            if ($rowNum === -1) {
                $rowNum = $row;
            }
            elseif ($rowNum !== $row) {
                return true;
            }
            $rowData[$col . $row] = $val;

            return null;
        }, [], null, $styleIdxInclude);

        return $rowData;
    }

    /**
     * Set read area and returns cell values of 1st row as array [address => value]
     *
     * Like readCellsFrom(), this method takes no column keys, because the result is keyed by cell address and renaming a column would corrupt it.
     *
     * @param string $areaRange
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readFirstRowCellsFrom(string $areaRange, ?bool $styleIdxInclude = null): array
    {
        return $this->setReadArea($areaRange)->readFirstRowCells($styleIdxInclude);
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [col][row]
     *      ['A' => [1 => _value_A1_], [2 => _value_A2_]],
     *      ['B' => [1 => _value_B1_], [2 => _value_B2_]]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readColumns($columnKeys = null, ?int $resultMode = null, ?bool $styleIdxInclude = null): array
    {
        if (is_int($columnKeys) && $columnKeys > 1 && $resultMode === null) {
            $resultMode = $columnKeys | Excel::KEYS_RELATIVE;
            $columnKeys = $columnKeys & Excel::KEYS_FIRST_ROW;
        }
        else {
            $resultMode = $resultMode | Excel::KEYS_RELATIVE;
        }

        return $this->readRows($columnKeys, $resultMode | Excel::KEYS_SWAP, $styleIdxInclude);
    }

    /**
     * Set read area and returns cell values as a two-dimensional array from default sheet [col][row]
     *
     * @param string $areaRange
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readColumnsFrom(string $areaRange, $columnKeys = null, ?int $resultMode = null, ?bool $styleIdxInclude = null): array
    {
        return $this->setReadArea($areaRange)->readColumns($columnKeys, $resultMode, $styleIdxInclude);
    }

    /**
     * Returns cell values and styles as a two-dimensional array [column][row]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readColumnsWithStyles($columnKeys = null, ?int $resultMode = null): array
    {
        $data = $this->readColumns($columnKeys, $resultMode, true);

        foreach ($data as $col => $colData) {
            foreach ($colData as $row => $cellData) {
                if (isset($cellData['s'])) {
                    $data[$col][$row]['s'] = $this->excel->getCompleteStyleByIdx($cellData['s']);
                }
            }
        }

        return $data;
    }

    /**
     * Set read area and returns cell values and styles as a two-dimensional array [column][row]
     *
     * @param string $areaRange
     * @param $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readColumnsWithStylesFrom(string $areaRange, $columnKeys = null, ?int $resultMode = null): array
    {
        return $this->setReadArea($areaRange)->readColumnsWithStyles($columnKeys, $resultMode);
    }

    /**
     * Returns values and styles of cells as array
     *
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readCells(?bool $styleIdxInclude = null): array
    {
        $data = [];
        $this->readCallback(static function($row, $col, $val) use (&$data) {
            $data[$col . $row] = $val;
        }, [], null, $styleIdxInclude);

        return $data;
    }

    /**
     * Set read area and returns cell values as a one-dimensional array [address => value]
     *
     * @param string $areaRange
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readCellsFrom(string $areaRange, ?bool $styleIdxInclude = null): array
    {
        return $this->setReadArea($areaRange)->readCells($styleIdxInclude);
    }

    /**
     * Returns cell values and styles as a one-dimensional array [address => value]:
     *      'v' => _value_
     *      's' => _styles_
     *      'f' => _formula_
     *      't' => _type_
     *      'o' => _original_value_
     *
     * @param string|null $styleKey If specified, only this style property will be returned (e.g. 'fill-color')
     *
     * @return array
     */
    public function readCellsWithStyles(?string $styleKey = null): array
    {
        $data = $this->readCells(true);
        foreach ($data as $cell => $cellData) {
            if (isset($cellData['s'])) {
                if ($styleKey) {
                    // properties such as 'fill-color' live inside a group, so the
                    // lookup has to happen on the flattened style
                    $flat = $this->excel->getCompleteStyleByIdx($cellData['s'], true);
                    $data[$cell]['s'] = array_key_exists($styleKey, $flat)
                        ? [$styleKey => $flat[$styleKey]]
                        : $this->excel->getCompleteStyleByIdx($cellData['s']);
                }
                else {
                    $data[$cell]['s'] = $this->excel->getCompleteStyleByIdx($cellData['s']);
                }
            }
        }

        return $data;
    }

    /**
     * Set read area and returns cell values and styles as a one-dimensional array [address => value]
     *
     * @param string $areaRange
     * @param string|null $styleKey If specified, only this style property will be returned (e.g. 'fill-color')
     *
     * @return array
     */
    public function readCellsWithStylesFrom(string $areaRange, ?string $styleKey = null): array
    {
        return $this->setReadArea($areaRange)->readCellsWithStyles($styleKey);
    }

    /**
     * Returns styles of cells as array
     *
     * @param bool|null $flat
     * @param string|null $part
     *
     * @return array
     */
    public function readCellStyles(?bool $flat = false, ?string $part = null): array
    {
        $cells = $this->readCells(true);
        $result = [];
        if ($part) {
            $flat = false;
        }
        foreach ($cells as $cell => $cellData) {
            if (isset($cellData['s'])) {
                $style = $this->excel->getCompleteStyleByIdx($cellData['s'], $flat);
                if ($cellData['t'] === 'date') {
                    //$style['format']['format-category'] = 'date';
                }
                $result[$cell] = $part ? ($style[$part] ?? []) : $style;
            }
            else {
                $result[$cell] = [];
            }
        }

        return $result;
    }

    /**
     * Reads cell values and passes them to a callback function
     *
     * @param callback $callback Callback function($row, $col, $value)
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     */
    public function readCallback(callable $callback, $columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null)
    {
        foreach ($this->nextRow($columnKeys, $resultMode, $styleIdxInclude) as $row => $rowData) {
            if (isset($rowData['__cells'], $rowData['__row'])) {
                $rowData = $rowData['__cells'];
            }
            foreach ($rowData as $col => $val) {
                if (empty($this->area['col_keys']) || array_key_exists($col, $this->area['col_keys']) || array_key_exists($col, $this->area['col_names'])) {
                    $needBreak = $callback($row, $col, $val);
                    if ($needBreak) {
                        return;
                    }
                }
            }
        }
    }

    protected function _rowTemplate(array $rowTemplate, array $columnKeys): array
    {
        $rowData = [];
        foreach (array_keys($rowTemplate) as $key) {
            if (isset($columnKeys[$key])) {
                $rowData[$columnKeys[$key]] = null;
                $this->area['col_names'][$columnKeys[$key]] = $key;
            }
            else {
                $rowData[$key] = null;
            }
        }

        return $rowData;
    }

    /**
     * Read cell values row by row, returns either an array of values or an array of arrays
     *
     *      nextRow(..., ...) : <rowNum> => [<colNum1> => <value1>, <colNum2> => <value2>, ...]
     *      nextRow(..., ..., true) : <rowNum> => [<colNum1> => ['v' => <value1>, 's' => <style1>], <colNum2> => ['v' => <value2>, 's' => <style2>], ...]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     * @param int|null $rowLimit
     *
     * @return \Generator|null
     */
    public function nextRow($columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null, ?int $rowLimit = 0): ?\Generator
    {
        // <dimension ref="A1:C1"/>
        // sometimes sheets doesn't contain this tag
        $this->dimension();

        if (!$columnKeys && is_int($resultMode) && ($resultMode & Excel::KEYS_FIRST_ROW)) {
            $firstRowValues = $this->readFirstRow();
            $columnKeys = array_keys($firstRowValues);
        }
        $readArea = $this->area;
        $rowTemplate = $readArea['col_keys'];
        if (!empty($columnKeys) && is_array($columnKeys)) {
            $firstRowKeys = is_int($resultMode) && ($resultMode & Excel::KEYS_FIRST_ROW);
            $columnKeys = array_combine(array_map('strtoupper', array_keys($columnKeys)), array_values($columnKeys));
        }
        elseif ($columnKeys === true || $columnKeys === false || $columnKeys === null) {
            $firstRowKeys = !!$columnKeys;
            $columnKeys = [];
        }
        elseif ($resultMode & Excel::KEYS_FIRST_ROW) {
            $firstRowKeys = true;
        }
        else {
            $firstRowKeys = !empty($readArea['first_row_keys']);
        }

        if ($columnKeys && ($resultMode & Excel::KEYS_FIRST_ROW)) {
            foreach ($this->nextRow([], 0, null, 1) as $firstRowData) {
                $columnKeys = array_merge($firstRowData, $columnKeys);
                break;
            }
        }
        $this->readRowNum = $this->countReadRows = 0;

        // When the read area restricts columns and numeric column keys are requested,
        // the template taken from col_keys is keyed by letter while the cell loop below
        // writes numeric keys. Left as is, every row would carry both halves and
        // readCallback() would keep the empty one. Transform the template, and anchor
        // the column offset to the area, so that the two sides agree.
        $colOffset = null;
        if ($rowTemplate && is_int($resultMode) && ($resultMode & (Excel::KEYS_COL_ZERO_BASED | Excel::KEYS_COL_ONE_BASED))) {
            $colOffset = $readArea['col_min'] - 1;
            if ($resultMode & Excel::KEYS_COL_ZERO_BASED) {
                $colOffset += 1;
            }
            $numericTemplate = [];
            foreach (array_keys($rowTemplate) as $colLetter) {
                if (is_array($columnKeys) && isset($columnKeys[$colLetter])) {
                    // an explicit name wins over the numeric key, and _rowTemplate()
                    // needs the letter to apply it
                    $numericTemplate[$colLetter] = null;
                    continue;
                }
                $colKey = Excel::colNum((string)$colLetter) - $colOffset;
                $numericTemplate[$colKey] = null;
                // makes readCallback() recognise the transformed key as part of the area
                $this->area['col_names'][$colKey] = $colLetter;
            }
            $rowTemplate = $numericTemplate;
        }

        // mapping col keys to col names
        $rowData = $rowTemplate = $this->_rowTemplate($rowTemplate, $columnKeys);

        $rowOffset = null;
        $rowCnt = -1;
        $rowMode = (bool)($resultMode & Excel::RESULT_MODE_ROW);

        foreach ($this->rawRows($readArea, (int)$rowLimit, $rowMode) as $rowNum => $rawRow) {
            $rowCnt++;
            $this->readRowNum = $rowNum;

            if ($rowOffset === null) {
                $rowOffset = 0;
                if (is_int($resultMode) && $resultMode) {
                    if ($resultMode & Excel::KEYS_ROW_ZERO_BASED) {
                        $rowOffset = $rowNum + ($firstRowKeys ? 1 : 0);
                    }
                    elseif ($resultMode & Excel::KEYS_ROW_ONE_BASED) {
                        $rowOffset = $rowNum - 1 + ($firstRowKeys ? 1 : 0);
                    }
                }
            }

            $rowData = $rowTemplate;
            foreach ($rawRow['cells'] as $colLetter => $cellData) {
                $colNum = Excel::colNum((string)$colLetter);

                // must be recorded after the column filter, otherwise the first
                // cell of the row is reported even when it lies outside the area
                if (!isset($this->area['first_row'])) {
                    $this->area['first_row'] = $rowNum;
                    $this->area['first_col'] = $colLetter;
                }
                if ($colOffset === null) {
                    $colOffset = $colNum - 1;
                    if (is_int($resultMode) && ($resultMode & Excel::KEYS_COL_ZERO_BASED)) {
                        $colOffset += 1;
                    }
                }
                if ($resultMode && ($resultMode & (Excel::KEYS_COL_ZERO_BASED | Excel::KEYS_COL_ONE_BASED))) {
                    $col = $colNum - $colOffset;
                }
                else {
                    $col = $colLetter;
                }
                if (is_array($columnKeys) && isset($columnKeys[$colLetter])) {
                    $col = $columnKeys[$colLetter];
                }

                if ($styleIdxInclude) {
                    $rowData[$col] = $cellData;
                }
                else {
                    $value = $cellData['v'];
                    if (is_string($value) && ($resultMode & Excel::TRIM_STRINGS)) {
                        $value = trim($value);
                    }
                    if (!($value === '' && ($resultMode & Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL))) {
                        $rowData[$col] = $value;
                    }
                }
            }

            // the first row of the area supplies the keys and is not returned itself
            if ($rowCnt === 0 && $firstRowKeys) {
                if (!$columnKeys) {
                    if ($styleIdxInclude) {
                        $columnKeys = array_combine(array_keys($rowData), array_column($rowData, 'v'));
                    }
                    else {
                        $columnKeys = $rowData;
                    }
                    $rowData = $rowTemplate = $this->_rowTemplate($rowData, $columnKeys);
                }
                continue;
            }

            if ($rowMode) {
                $rowData = [
                    '__cells' => $rowData,
                    '__row' => $rawRow['attrs'],
                ];
            }

            yield ($rowNum - $rowOffset) => $rowData;
        }

        return null;
    }

    /**
     * Reset read generator
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     * @param int|null $rowLimit
     *
     * @return \Generator|null
     */
    public function reset($columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null, ?int $rowLimit = 0): ?\Generator
    {
        $this->generator = $this->nextRow($columnKeys, $resultMode, $styleIdxInclude, $rowLimit);
        $this->countReadRows = 0;

        return $this->generator;
    }

    /**
     * Rewind read generator, alias of reset()
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     * @param int|null $rowLimit
     *
     * @return \Generator|null
     */
    public function rewind($columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null, ?int $rowLimit = 0): ?\Generator
    {

        return $this->reset($columnKeys, $resultMode, $styleIdxInclude, $rowLimit);
    }

    /**
     * Read the next row from the generator
     *
     * @return mixed
     */
    public function readNextRow()
    {
        if (!$this->generator) {
            $this->reset();
        }
        if ($this->countReadRows > 0) {
            $this->generator->next();
        }
        if ($result = $this->generator->current()) {
            $this->countReadRows++;
        }

        return $result;
    }

    /**
     * Get the number of the last row read
     *
     * @return int
     */
    public function getReadRowNum(): int
    {
        return $this->readRowNum;
    }

    /**
     * Returns true if the cell is merged
     *
     * @param string $cellAddress
     *
     * @return bool
     */
    public function isMerged(string $cellAddress): bool
    {
        foreach ($this->getMergedCells() as $range) {
            if (Helper::inRange($cellAddress, $range)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Get merged range for the cell
     *
     * @param string $cellAddress
     *
     * @return string|null
     */
    public function mergedRange(string $cellAddress): ?string
    {
        foreach ($this->getMergedCells() as $range) {
            if (Helper::inRange($cellAddress, $range)) {
                return $range;
            }
        }

        return null;
    }

}
