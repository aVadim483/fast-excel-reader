<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceBookReader;

/**
 * Format independent part of a workbook reader
 *
 * Holds everything that does not depend on how the workbook is stored: the
 * sheet collection and its selectors, the delegating read* helpers, date and
 * number format handling, and style composition. A concrete reader supplies
 * _prepare(), which fills the sheet list, shared strings and styles.
 *
 * @see Excel for the XLSX implementation
 */
abstract class AbstractBook implements InterfaceBookReader
{
    /** @var int Maximum number of rows in Excel 2007+ */
    public const EXCEL_2007_MAX_ROW = 1048576;
    /** @var int Maximum number of columns in Excel 2007+ */
    public const EXCEL_2007_MAX_COL = 16384;

    /** @var int Use original keys (column names like 'A', 'B', etc.) */
    public const KEYS_ORIGINAL = 0;
    /** @var int Use the first row values as keys */
    public const KEYS_FIRST_ROW = 1;
    /** @var int Use zero-based row indices as keys */
    public const KEYS_ROW_ZERO_BASED = 2;
    /** @var int Use zero-based column indices as keys */
    public const KEYS_COL_ZERO_BASED = 4;
    /** @var int Use both zero-based row and column indices as keys */
    public const KEYS_ZERO_BASED = 6; // KEYS_ROW_ZERO_BASED & KEYS_COL_ZERO_BASED
    /** @var int Use one-based row indices as keys */
    public const KEYS_ROW_ONE_BASED = 8;
    /** @var int Use one-based column indices as keys */
    public const KEYS_COL_ONE_BASED = 16;
    /** @var int Use both one-based row and column indices as keys */
    public const KEYS_ONE_BASED = 24; // KEYS_ROW_ONE_BASED & KEYS_COL_ONE_BASED
    /** @var int Use relative keys */
    public const KEYS_RELATIVE = 32;
    /** @var int Swap keys and values */
    public const KEYS_SWAP = 64;

    /**
     * nextRow() returns cells & row attributes
     * ['__cells' => [...], '__row' => [...]]
     */
    public const RESULT_MODE_ROW = 1024;

    /** @var int Trim string values */
    public const TRIM_STRINGS = 2048;
    /** @var int Treat empty strings as empty cells */
    public const TREAT_EMPTY_STRING_AS_EMPTY_CELL = 4096;

    protected string $file;

    protected array $sharedStrings = [];

    protected array $styles = [];

    /** @var AbstractSheet[] */
    protected array $sheets = [];

    protected int $defaultSheetId;

    protected ?string $dateFormat = null;

    /** @var \Closure|callable|bool|null  */
    protected $dateFormatter = null;

    protected bool $date1904 = false;

    protected string $timezone;

    protected array $builtinFormats = [];

    protected array $names = [];

    /**
     * Read the workbook structure: sheets, shared strings, styles, defined names
     *
     * @param string $file
     *
     * @return void
     */
    abstract protected function _prepare(string $file): void;

    /**
     * Set directory for temporary files
     *
     * @param string $tempDir
     *
     * @return void
     */
    abstract public static function setTempDir($tempDir);

    /**
     * Fill $this->styles['_'] with the complete style tables
     *
     * The expected shape is the one getCompleteStyleByIdx() composes from:
     * 'numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs' and 'cellXfs'.
     *
     * @return void
     */
    abstract protected function _loadCompleteStyles();

    /**
     * Excel constructor
     *
     * @param string|null $file
     * @param string|null $tempDir
     */
    public function __construct(?string $file = null, ?string $tempDir = '')
    {
        $this->builtinFormats = [
            0 => ['pattern' => 'General', 'category' => 'general'],
            1 => ['pattern' => '0', 'category' => 'number'],
            2 => ['pattern' => '0.00', 'category' => 'number'],
            3 => ['pattern' => '#,##0', 'category' => 'number'],
            4 => ['pattern' => '#,##0.00', 'category' => 'number'],
            9 => ['pattern' => '0%', 'category' => 'number'],
            10 => ['pattern' => '0.00%', 'category' => 'number'],
            11 => ['pattern' => '0.00E+00', 'category' => 'number'],
            12 => ['pattern' => '# ?/?', 'category' => 'general'],
            13 => ['pattern' => '# ??/??', 'category' => 'general'],
            14 => ['pattern' => 'mm-dd-yy', 'category' => 'date'], // Short date
            15 => ['pattern' => 'd-mmm-yy', 'category' => 'date'],
            16 => ['pattern' => 'd-mmm', 'category' => 'date'],
            17 => ['pattern' => 'mmm-yy', 'category' => 'date'],
            18 => ['pattern' => 'h:mm AM/PM', 'category' => 'date'],
            19 => ['pattern' => 'h:mm:ss AM/PM', 'category' => 'date'],
            20 => ['pattern' => 'h:mm', 'category' => 'date'], // Short time
            21 => ['pattern' => 'h:mm:ss', 'category' => 'date'], // Long time
            22 => ['pattern' => 'm/d/yy h:mm', 'category' => 'date'], // Date-time
            37 => ['pattern' => '#,##0 ;(#,##0)', 'category' => 'number'],
            38 => ['pattern' => '#,##0 ;[Red](#,##0)', 'category' => 'number'],
            39 => ['pattern' => '#,##0.00;(#,##0.00)', 'category' => 'number'],
            40 => ['pattern' => '#,##0.00;[Red](#,##0.00)', 'category' => 'number'],
            45 => ['pattern' => 'mm:ss', 'category' => 'date'],
            46 => ['pattern' => '[h]:mm:ss', 'category' => 'date'],
            47 => ['pattern' => 'mmss.0', 'category' => 'date'],
            48 => ['pattern' => '##0.0E+0', 'category' => 'number'],
            49 => ['pattern' => '@', 'category' => 'string'],
        ];

        if (class_exists('IntlDateFormatter', false)) {
            $formatter = new \IntlDateFormatter(null, \IntlDateFormatter::SHORT, \IntlDateFormatter::NONE);
            $pattern = $formatter->getPattern();
            $this->builtinFormats[14]['pattern'] = str_replace('#', 'yy', str_replace(['M', 'y'], ['m', 'yyyy'], str_replace('yy', '#', $pattern)));
            if (preg_match('/([^a-z])/i', $pattern, $m)) {
                $dateDelim = $m[1];
                $this->builtinFormats[15]['pattern'] = str_replace('-', $dateDelim, $this->builtinFormats[15]['pattern']);
                $this->builtinFormats[16]['pattern'] = str_replace('-', $dateDelim, $this->builtinFormats[16]['pattern']);
                $this->builtinFormats[17]['pattern'] = str_replace('-', $dateDelim, $this->builtinFormats[17]['pattern']);
            }

            $formatter = new \IntlDateFormatter(null, \IntlDateFormatter::NONE, \IntlDateFormatter::SHORT);
            $this->builtinFormats[20]['pattern'] = str_replace('HH', 'h', $formatter->getPattern());

            $formatter = new \IntlDateFormatter(null, \IntlDateFormatter::NONE, \IntlDateFormatter::MEDIUM);
            $this->builtinFormats[21]['pattern'] = str_replace('HH', 'h', $formatter->getPattern());

            $this->builtinFormats[22]['pattern'] = $this->builtinFormats[14]['pattern'] . ' ' . $this->builtinFormats[20]['pattern'];
        }

        $this->timezone = date_default_timezone_get();
        $this->dateFormatter = function ($value, $format = null) {
            if ($format || $this->dateFormat) {
                return gmdate($format ?: $this->dateFormat, $value);
            }
            return $value;
        };

        if (!empty($tempDir)) {
            static::setTempDir($tempDir);
        }

        if ($file) {
            $this->file = $file;
            $this->_prepare($file);
        }
    }

    /**
     * @param int|null $numFmtId
     * @param string $pattern
     *
     * @return bool
     */
    protected function _isDatePattern(?int $numFmtId, string $pattern): bool
    {
        if ($numFmtId && (
            ($numFmtId >= 14 && $numFmtId <= 22)
            || ($numFmtId >= 45 && $numFmtId <= 47)
            || ($numFmtId >= 27 && $numFmtId <= 36)
            || ($numFmtId >= 50 && $numFmtId <= 58)
            || ($numFmtId >= 71 && $numFmtId <= 81)
            )) {
            return true;
        }
        if ($pattern) {
            if (preg_match('/^\[\$-[0-9A-F]{3,4}].+/', $pattern)) {
                return true;
            }
            return (bool)preg_match('/yy|mm|dd|h|MM|ss|[\/\.][dm](;.+)?/', $pattern);
        }

        return false;
    }

    /**
     * @param int|null $numFmtId
     * @param string $pattern
     *
     * @return bool
     */
    protected function _isNumberPattern(?int $numFmtId, string $pattern): bool
    {
        if (preg_match('/^0+(\.0+)?$/', $pattern)) {
            return true;
        }

        return false;
    }

    /**
     * Converts an alphabetic column index to a numeric
     *
     * @param string $colLetter
     *
     * @return int
     */
    public static function colNum(string $colLetter): int
    {

        return Helper::colNumber($colLetter);
    }

    /**
     * Convert column number to letter
     *
     * @param int $colNumber ONE based
     *
     * @return string
     */
    public static function colLetter(int $colNumber): string
    {

        return Helper::colLetter($colNumber);
    }

    /**
     * Convert date to timestamp
     *
     * @param $excelDateTime
     *
     * @return int
     */
    public function timestamp($excelDateTime): int
    {
        $excelDateTime = trim($excelDateTime);
        if (is_numeric($excelDateTime)) {
            $d = floor($excelDateTime);
            $t = $excelDateTime - $d;
            if ($this->date1904) {
                $d += 1462; // days since 1904
            }

            // Adjust for Excel erroneously treating 1900 as a leap year.
            if ($d <= 59) {
                $d++;
            }
            $t = (abs($d) > 0) ? ($d - 25569) * 86400 + round($t * 86400) : round($t * 86400);
        }
        elseif (preg_match('/^[\d\.\-\/:\s]+$/', $excelDateTime)) {
            if ($this->timezone !== 'UTC') {
                date_default_timezone_set('UTC');
            }
            $t = strtotime($excelDateTime);
            if ($this->timezone !== 'UTC') {
                date_default_timezone_set($this->timezone);
            }
        }
        else {
            // string is not a date
            $t = 0;
        }

        return (int)$t;
    }

    /**
     * Set date format for reading
     *
     * @param string $dateFormat
     *
     * @return $this
     */
    public function setDateFormat(string $dateFormat): AbstractBook
    {
        $this->dateFormat = $dateFormat;

        return $this;
    }

    /**
     * Get current date format
     *
     * @return string|null
     */
    public function getDateFormat(): ?string
    {
        return $this->dateFormat;
    }

    /**
     * Format date value
     *
     * @param mixed $value
     * @param string|null $format
     * @param int|null $styleIdx
     *
     * @return false|mixed|string
     */
    public function formatDate($value, $format = null, $styleIdx = null)
    {
        if ($this->dateFormatter && $this->dateFormatter !== true) {
            return ($this->dateFormatter)($value, $format, $styleIdx);
        }

        return $value;
    }

    /**
     * Set custom date formatter
     *
     * @param \Closure|callable|string|bool|null $formatter
     *
     * @return $this
     */
    public function dateFormatter($formatter): AbstractBook
    {
        if ($formatter === false || $formatter === null) {
            $this->dateFormatter = $formatter;
        }
        elseif ($formatter === true) {
            $this->dateFormatter = function ($value, $format = null, $styleIdx = null) {
                if ($styleIdx !== null && $pattern = $this->getDateFormatPattern($styleIdx)) {
                    return gmdate($pattern, $value);
                }
                elseif ($format || $this->dateFormat) {
                    return gmdate($format ?: $this->dateFormat, $value);
                }
                return $value;
            };
        }
        elseif (is_string($formatter)) {
            $this->dateFormat = $formatter;
            $this->dateFormatter = function ($value, $format = null) {
                if ($format || $this->dateFormat) {
                    return gmdate($format ?: $this->dateFormat, $value);
                }
                return $value;
            };
        }
        else {
            $this->dateFormatter = $formatter;
        }

        return $this;
    }

    /**
     * Get date formatter
     *
     * @return callable|\Closure|bool|null
     */
    public function getDateFormatter()
    {
        return $this->dateFormatter;
    }

    /**
     * Get style array by style index
     *
     * @param int $styleIdx
     *
     * @return array
     */
    public function styleByIdx($styleIdx): array
    {
        return $this->styles['cellXfs'][$styleIdx] ?? [];
    }

    /**
     * Get string by index
     *
     * @param int $stringId
     *
     * @return string|null
     */
    public function sharedString($stringId): ?string
    {
        return $this->sharedStrings[$stringId] ?? null;
    }

    /**
     * Get defined names of workbook
     *
     * @return array
     */
    public function getDefinedNames(): array
    {
        return $this->names;
    }

    /**
     * Get names array of all sheets
     *
     * @return array
     */
    public function getSheetNames(): array
    {
        $result = [];
        foreach ($this->sheets as $sheetId => $sheet) {
            $result[$sheetId] = $sheet->name();
        }
        return $result;
    }

    /**
     * Get current or specified sheet
     *
     * @param string|null $name
     *
     * @return AbstractSheet|null
     */
    public function sheet(?string $name = null): ?AbstractSheet
    {
        $resultId = null;
        if (!$name) {
            $resultId = $this->defaultSheetId;
        }
        else {
            foreach ($this->sheets as $sheetId => $sheet) {
                if ($sheet->isName($name)) {
                    $resultId = $sheetId;
                    break;
                }
            }
        }
        if ($resultId && isset($this->sheets[$resultId])) {
            return $this->sheets[$resultId];
        }

        return null;
    }

    /**
     * Get sheet object by name and optionally set read area and options
     *
     * @param string|null $name
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return AbstractSheet
     */
    public function getSheet(?string $name = null, ?string $areaRange = null, ?bool $firstRowKeys = false): AbstractSheet
    {
        $sheet = null;
        if (!$name) {
            $sheet = $this->sheet();
        }
        else {
            foreach ($this->sheets as $foundSheet) {
                if ($foundSheet->isName($name)) {
                    $sheet = $foundSheet;
                    break;
                }
            }
            if (!$sheet) {
                throw new Exception('Sheet name "' . $name . '" not found');
            }
        }

        if ($areaRange) {
            $sheet->setReadArea($areaRange, $firstRowKeys);
        }

        return $sheet;
    }

    /**
     * Returns a sheet by ID
     *
     * @param int $sheetId
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return AbstractSheet
     */
    public function getSheetById(int $sheetId, ?string $areaRange = null, ?bool $firstRowKeys = false): AbstractSheet
    {
        if (!isset($this->sheets[$sheetId])) {
            throw new Exception('Sheet ID "' . $sheetId . '" not found');
        }
        if ($areaRange) {
            $this->sheets[$sheetId]->setReadArea($areaRange, $firstRowKeys);
        }

        return $this->sheets[$sheetId];
    }

    /**
     * Returns the first sheet as default
     *
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return AbstractSheet
     */
    public function getFirstSheet(?string $areaRange = null, ?bool $firstRowKeys = false): AbstractSheet
    {
        $sheetId = array_key_first($this->sheets);
        $sheet = $this->sheets[$sheetId];
        if ($areaRange) {
            $sheet->setReadArea($areaRange, $firstRowKeys);
        }

        return $sheet;
    }

    /**
     * Selects default sheet by name
     *
     * @param string $name
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return AbstractSheet
     */
    public function selectSheet(string $name, ?string $areaRange = null, ?bool $firstRowKeys = false): AbstractSheet
    {
        $sheet = $this->getSheet($name, $areaRange, $firstRowKeys);
        $this->defaultSheetId = $sheet->id();

        return $sheet;
    }

    /**
     * Selects default sheet by ID
     *
     * @param int $sheetId
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return AbstractSheet
     */
    public function selectSheetById(int $sheetId, ?string $areaRange = null, ?bool $firstRowKeys = false): AbstractSheet
    {
        $sheet = $this->getSheetById($sheetId, $areaRange, $firstRowKeys);
        $this->defaultSheetId = $sheet->id();

        return $sheet;
    }

    /**
     * Selects the first sheet as default
     *
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return AbstractSheet
     */
    public function selectFirstSheet(?string $areaRange = null, ?bool $firstRowKeys = false): AbstractSheet
    {
        $sheet = $this->getFirstSheet($areaRange, $firstRowKeys);
        $this->defaultSheetId = $sheet->id();

        return $sheet;
    }

    /**
     * Array of all sheets
     *
     * @return AbstractSheet[]
     */
    public function sheets(): array
    {
        return $this->sheets;
    }

    /**
     * Array of visible sheets only
     *
     * @return AbstractSheet[]
     */
    public function visibleSheets(): array
    {
        $result = [];
        foreach ($this->sheets as $sheetId => $sheet) {
            if ($sheet->isVisible()) {
                $result[$sheetId] = $sheet;
            }
        }

        return $result;
    }

    /**
     * Array of hidden sheets only
     *
     * @return AbstractSheet[]
     */
    public function hiddenSheets(): array
    {
        $result = [];
        foreach ($this->sheets as $sheetId => $sheet) {
            if ($sheet->isHidden()) {
                $result[$sheetId] = $sheet;
            }
        }

        return $result;
    }

    /**
     * Returns TRUE if a sheet with the given name exists
     *
     * @param string $name
     *
     * @return bool
     */
    public function sheetExists(string $name): bool
    {
        foreach ($this->sheets as $sheet) {
            if ($sheet->isName($name)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Returns the number of sheets in the workbook
     *
     * @return int
     */
    public function countSheets(): int
    {
        return count($this->sheets);
    }

    /**
     * Returns statistics of the workbook: per-sheet breakdown and totals
     *
     * [
     *      'sheets' => [
     *          '<sheetName>' => ['rows' => [...], 'cols' => [...], 'cells' => ['total' => int, 'filled' => int]],
     *          ...
     *      ],
     *      'total' => [
     *          'sheets'  => int,   // number of sheets
     *          'visible' => int,   // number of visible sheets
     *          'hidden'  => int,   // number of hidden sheets
     *          'rows'    => int,   // sum of actual rows over all sheets
     *          'cells'   => ['total' => int, 'filled' => int],
     *      ],
     * ]
     *
     * Note: scans every sheet fully (see Sheet::stat()); expensive on large workbooks.
     *
     * @return array
     */
    public function stat(): array
    {
        $sheets = [];
        $totalRows = $totalCells = $filledCells = 0;
        foreach ($this->sheets as $sheet) {
            $stat = $sheet->stat();
            $sheets[$sheet->name()] = $stat;
            $totalRows += $stat['rows']['count'] ?? 0;
            $totalCells += $stat['cells']['total'] ?? 0;
            $filledCells += $stat['cells']['filled'] ?? 0;
        }

        return [
            'sheets' => $sheets,
            'total' => [
                'sheets' => $this->countSheets(),
                'visible' => count($this->visibleSheets()),
                'hidden' => count($this->hiddenSheets()),
                'rows' => $totalRows,
                'cells' => ['total' => $totalCells, 'filled' => $filledCells],
            ],
        ];
    }

    /**
     * Set top left and right bottom of read area
     *
     * @param string $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return AbstractSheet
     */
    public function setReadArea(string $areaRange, ?bool $firstRowKeys = false): AbstractSheet
    {
        $sheet = $this->sheets[$this->defaultSheetId];
        if (preg_match('/^\w+$/', $areaRange)) {
            foreach ($this->getDefinedNames() as $name => $range) {
                if ($name === $areaRange) {
                    [$sheetName, $definedRange] = explode('!', $range);
                    $sheet = $this->selectSheet($sheetName);
                    $areaRange = $definedRange;
                    break;
                }
            }
        }

        return $sheet->setReadArea($areaRange, $firstRowKeys);
    }

    /**
     * Set top left of read area
     *
     * @param string $topLeftCell
     * @param bool|null $firstRowKeys
     *
     * @return AbstractSheet
     */
    public function from(string $topLeftCell, ?bool $firstRowKeys = false): AbstractSheet
    {
        return $this->setReadArea($topLeftCell, $firstRowKeys);
    }

    /**
     * Reads cell values and passes them to a callback function
     *
     * @param callback $callback
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     */
    public function readCallback(callable $callback, ?int $resultMode = null, ?bool $styleIdxInclude = null)
    {
        $this->sheets[$this->defaultSheetId]->readCallback($callback, [], $resultMode, $styleIdxInclude);
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [row][col]
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
        return $this->sheets[$this->defaultSheetId]->readRows($columnKeys, $resultMode, $styleIdxInclude);
    }

    /**
     * Returns cell values and styles as a two-dimensional array from default sheet [row][col]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readRowsWithStyles($columnKeys = [], ?int $resultMode = null): array
    {
        return $this->sheets[$this->defaultSheetId]->readRowsWithStyles($columnKeys, $resultMode);
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [col][row]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readColumns($columnKeys = null, ?int $resultMode = null): array
    {
        return $this->sheets[$this->defaultSheetId]->readColumns($columnKeys, $resultMode);
    }

    /**
     * Returns cell values and styles as a two-dimensional array from default sheet [col][row]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readColumnsWithStyles($columnKeys = null, ?int $resultMode = null): array
    {
        return $this->sheets[$this->defaultSheetId]->readColumnsWithStyles($columnKeys, $resultMode);
    }

    /**
     * Returns the values of all cells as array
     *
     * @return array
     */
    public function readCells(): array
    {
        return $this->sheets[$this->defaultSheetId]->readCells();
    }

    /**
     * Returns the values and styles of all cells as array
     *
     * @return array
     */
    public function readCellsWithStyles(): array
    {
        return $this->sheets[$this->defaultSheetId]->readCellsWithStyles();
    }

    /**
     * Returns the styles of all cells as array
     *
     * @param bool|null $flat
     *
     * @return array
     */
    public function readCellStyles(?bool $flat = false): array
    {
        return $this->sheets[$this->defaultSheetId]->readCellStyles($flat);
    }

    /**
     * Read all workbook styles
     *
     * @return array
     */
    public function readStyles(): array
    {
        if (!isset($this->styles['_'])) {
            $this->styles['_'] = [];
            $this->_loadCompleteStyles();
        }

        return $this->styles['_'];
    }

    /**
     * Get complete style by style index
     *
     * @param int $styleIdx
     * @param bool|null $flat
     *
     * @return array
     */
    public function getCompleteStyleByIdx(int $styleIdx, ?bool $flat = false): array
    {
        static $completedStyles = [];

        if (![$this->file]) {
            return [];
        }

        if (!isset($completedStyles[$this->file][$styleIdx])) {
            if ($styleIdx !== 0) {
                $result = $this->getCompleteStyleByIdx(0);
            }
            else {
                $result = [];
            }
            $styles = $this->readStyles();
            if (isset($styles['cellXfs'][$styleIdx])) {
                // Excel first takes the style settings with the xfId number from <cellStyleXfs>
                // and then applies the changes specified directly in <xf>
                $baseStyleId = $styles['cellXfs'][$styleIdx]['xfId'] ?? -1;
                if ($baseStyleId >= 0 && isset($styles['cellStyleXfs'][$baseStyleId])) {
                    $baseStyle = $styles['cellStyleXfs'][$baseStyleId];
                }
                else {
                    $baseStyle = [];
                }
                $result = array_replace_recursive($result, $baseStyle, $styles['cellXfs'][$styleIdx]);
                if (isset($result['xfId'])) {
                    unset($result['xfId']);
                }
            }

            if (isset($result['numFmtId']) && isset($styles['numFmts'][$result['numFmtId']])) {
                if (isset($result['format'])) {
                    $result['format'] = array_replace_recursive($result['format'], $styles['numFmts'][$result['numFmtId']]);
                }
                else {
                    $result['format'] = $styles['numFmts'][$result['numFmtId']];
                }
                unset($result['numFmtId']);
            }

            if (isset($result['fontId']) && isset($styles['fonts'][$result['fontId']])) {
                if (isset($result['font'])) {
                    $result['font'] = array_replace_recursive($result['font'], $styles['fonts'][$result['fontId']]);
                }
                else {
                    $result['font'] = $styles['fonts'][$result['fontId']];
                }
                unset($result['fontId']);
            }

            if (isset($result['fillId']) && isset($styles['fills'][$result['fillId']])) {
                if (isset($result['fill'])) {
                    $result['fill'] = array_replace_recursive($result['fill'], $styles['fills'][$result['fillId']]);
                }
                else {
                    $result['fill'] = $styles['fills'][$result['fillId']];
                }
                unset($result['fillId']);
            }

            if (isset($result['borderId']) && isset($styles['borders'][$result['borderId']])) {
                if (isset($result['border'])) {
                    $result['border'] = array_replace_recursive($result['border'], $styles['borders'][$result['borderId']]);
                }
                else {
                    $result['border'] = $styles['borders'][$result['borderId']];
                }
                unset($result['borderId']);
            }

            $completedStyles[$this->file][$styleIdx] = $result;
        }
        else {
            $result = $completedStyles[$this->file][$styleIdx];
        }

        if ($flat && $result) {
            $result = array_merge(...array_values($result));
        }

        return $result;
    }

    /**
     * Get format pattern by style index
     *
     * @param int $styleIdx
     *
     * @return mixed|string
     */
    public function getFormatPattern(int $styleIdx)
    {
        $style = $this->getCompleteStyleByIdx($styleIdx);

        return $style['format']['format-pattern'] ?? '';
    }

    /**
     * Convert Excel date format pattern to PHP date format pattern
     *
     * @param string $pattern
     *
     * @return string|null
     */
    public function _convertDateFormatPattern($pattern): ?string
    {
        static $patterns = [];

        if (isset($patterns[$pattern])) {
            return $patterns[$pattern];
        }

        if ($this->_isDatePattern(null, $pattern) && preg_match('/^(\[.+])?([^;]+)(;.*)?/', $pattern, $m)) {
            if (strpos($m[2], 'AM/PM')) {
                $am = true;
                $pattern = str_replace('AM/PM', 'A', $m[2]);
            }
            elseif (strpos($m[1], 'am/pm')) {
                $am = true;
                $pattern = str_replace('am/pm', 'a', $m[2]);
            }
            else {
                $am = false;
                $pattern = $m[2];
            }
            $pattern = str_replace(['\\ ', '\\-', '\\/'], [' ', '-', '/'], $pattern);
            $pattern = preg_replace(['/^mm(\W)s/', '/h(\W)mm$/', '/h(\W)mm([^m])/', '/([^m])mm(\W)s/'], ['i$1s', 'h$1i', 'h$1i$2', '$1i$1s'], $pattern);
            if ($am) {
                $pattern = str_replace(['hh', 'h'], ['h', 'g'], $pattern);
            }
            else {
                $pattern = str_replace(['hh', 'h'], ['H', 'G'], $pattern);
            }
            if (strpos($pattern, 'dd') !== false) {
                $pattern = str_replace('dd', 'd', $pattern);
            }
            else {
                $pattern = str_replace('d', 'j', $pattern);
            }
            $pattern = str_replace('mmmm', 'F', $pattern);
            $pattern = str_replace('mmm', 'M', $pattern);
            if (strpos($pattern, 'mm') !== false) {
                $pattern = str_replace('mm', 'm', $pattern);
            }
            else {
                $pattern = str_replace('m', 'n', $pattern);
            }
            $convert = [
                'ss' => 's',
                'dddd' => 'l',
                'ddd' => 'D',
                'mmmm' => 'F',
                'mmm' => 'M',
                'yyyy' => 'Y',
                'yy' => 'y',
            ];
            $patterns[$pattern] = str_replace(array_keys($convert), array_values($convert), $pattern);

            return $patterns[$pattern];
        }

        return null;
    }

    /**
     * Get PHP date format pattern by style index
     *
     * @param int $styleIdx
     *
     * @return string|null
     */
    public function getDateFormatPattern(int $styleIdx): ?string
    {
        $pattern = $this->getFormatPattern($styleIdx);
        if ($pattern) {
            return $this->_convertDateFormatPattern($pattern);
        }

        return null;
    }

}
