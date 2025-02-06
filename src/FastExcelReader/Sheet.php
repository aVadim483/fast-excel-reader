<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceBookReader;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;
use avadim\FastExcelReader\Interfaces\InterfaceXmlReader;

class Sheet implements InterfaceSheetReader
{
    public InterfaceBookReader $excel;

    protected string $zipFilename;

    protected string $sheetId;

    protected string $name;

    protected string $state = '';

    protected string $pathInZip;

    protected ?array $dimension = null;

    protected ?array $cols = null;

    protected ?bool $active = null;
    protected array $area = [];

    protected array $props = [];

    protected array $images = [];

    protected ?array $mergedCells = null;

    /** @var Reader */
    protected InterfaceXmlReader $xmlReader;

    protected int $readRowNum = 0;

    /** @var mixed */
    protected $preReadFunc = null;

    /** @var mixed */
    protected $postReadFunc = null;

    protected array $readNodeFunc = [];

    /**
     * @var \Generator|null
     */
    protected ?\Generator $generator = null;

    protected int $countReadRows = 0;

    protected array $sharedFormulas = [];

    protected int $countImages = -1; // -1 - unknown

    /**
     * @var array<array{
     *  type: string,
     *  sqref: string,
     *  formula1: ?string,
     *  formula2: ?string,
     * }>|null
     */
    protected ?array $validations = null;

    protected ?array $conditionals = null;

    protected ?array $rowHeights = null;

    protected ?array $colWidths = null;

    protected float $defaultRowHeight = 15.0;

    protected ?array $tabProperties = null;

    /**
     * @param string $sheetName
     * @param string $sheetId
     * @param string $file
     * @param string $path
     * @param $excel
     */
    public function __construct(string $sheetName, string $sheetId, string $file, string $path, $excel)
    {
        $this->excel = $excel;
        $this->name = $sheetName;
        $this->sheetId = $sheetId;
        $this->zipFilename = $file;
        $this->pathInZip = $path;

        $this->area = [
            'row_min' => 1,
            'col_min' => 1,
            'row_max' => Helper::EXCEL_2007_MAX_ROW,
            'col_max' => Helper::EXCEL_2007_MAX_COL,
            'first_row_keys' => false,
            'col_keys' => [],
        ];
    }

    /**
     * @param $cell
     * @param array|null $additionalData
     *
     * @return mixed
     */
    protected function _cellValue($cell, ?array &$additionalData = [])
    {
        // Determine data type and style index
        $attributeT = $dataType = (string)$cell->getAttribute('t');
        $styleIdx = (int)$cell->getAttribute('s');
        $address = $cell->attributes['r']->value;

        $cellValue = $formula = null;
        if ($cell->hasChildNodes()) {
            foreach($cell->childNodes as $node) {
                if ($node->nodeName === 'v') {
                    $cellValue = $node->nodeValue;
                    break;
                }
            }
            foreach($cell->childNodes as $node) {
                if ($node->nodeName === 'f') {
                    $formula = $this->_cellFormula($node, $address);
                    break;
                }
            }
            if ($cellValue === null) {
                $cellValue = $formula;
            }
        }
        elseif ($styleIdx) {
            $cellValue = '';
        }

        // Value is a shared string
        if ($dataType === 's') {
            if (is_numeric($cellValue) && null !== ($str = $this->excel->sharedString((int)$cellValue))) {
                $cellValue = $str;
            }
        }
        $formatCode = null;
        if (($cellValue !== null) && ($cellValue !== '') && ($dataType === '' || $dataType === 'n'  || $dataType === 's')) { // number or date as string
            if ($styleIdx > 0 && ($style = $this->excel->styleByIdx($styleIdx))) {
                if (isset($style['formatType'])) {
                    $dataType = $style['formatType'];
                }
                if (isset($style['format'])) {
                    $formatCode = $style['format'];
                }
            }
        }

        $originalValue = $cellValue;
        $value = '';

        switch ( $dataType ) {
            case 'b':
                // Value is boolean
                $value = (bool)$cellValue;
                $dataType = 'bool';
                break;

            case 'inlineStr':
                // Value is rich text inline
                $value = $cell->textContent;
                if ($value && $originalValue === null) {
                    $originalValue = $value;
                }
                $dataType = 'string';
                break;

            case 'e':
                // Value is an error message
                $value = (string)$cellValue;
                $dataType = 'error';
                break;

            case 'd':
            case 'date':
                if (($cellValue === null) || (trim($cellValue) === '')) {
                    $dataType = 'date';
                }
                elseif ($this->excel->getDateFormatter() === false) {
                    if ($attributeT !== 's' && is_numeric($cellValue)) {
                        $value = $this->excel->timestamp($cellValue);
                    }
                    else {
                        $value = $originalValue;
                    }
                }
                elseif (($timestamp = $this->excel->timestamp($cellValue))) {
                    // Value is a date and non-empty
                    $value = $this->excel->formatDate($timestamp, null, $styleIdx);
                    $dataType = 'date';
                }
                else {
                    // Value is not a date, load its original value
                    $value = (string)$cellValue;
                    //$dataType = 'string';
                }
                $dataType = 'date';
                break;

            default:
                if ($dataType === 'n' || $dataType === 'number') {
                    $dataType = 'number';
                }
                elseif ($dataType === 's' || $dataType === 'string') {
                    $dataType = 'string';
                }
                if ($cellValue === null) {
                    $value = null;
                }
                else {
                    // Value is a string
                    $value = (string)$cellValue;

                    // Check for numeric values
                    if ($dataType !== 'string' && is_numeric($value)) {
                        if (false !== $castedValue = filter_var($value, FILTER_VALIDATE_INT)) {
                            $value = $castedValue;
                            $dataType = 'number';
                        }
                        elseif (strlen($value) > 2 && !($value[0] === '0' && $value[1] !== '.') && false !== $castedValue = filter_var($value, FILTER_VALIDATE_FLOAT)) {
                            $value = $castedValue;
                            $dataType = 'number';
                        }
                        /*
                        if ($formatCode && preg_match('/\.(0+)$/', $formatCode, $m)) {
                            $value = round($value, strlen($m[1]));
                        }
                        */
                    }
                }
        }
        $additionalData = ['v' => $value, 's' => $styleIdx, 'f' => $formula, 't' => $dataType, 'o' => $originalValue];

        return $value;
    }

    /**
     * @param $node
     * @param string $address
     *
     * @return string
     */
    protected function _cellFormula($node, string $address): string
    {
        $shared = (string)$node->getAttribute('t') === 'shared';
        $si = (string)$node->getAttribute('si');
        $formula = $node->nodeValue;
        if ($formula) {
            if ($formula[0] !== '=') {
                $formula = '=' . $formula;
            }
            if ($shared && $si > '') {
                $this->sharedFormulas[$si] = $formula;
            }
        }
        elseif ($shared && $si > '' && isset($this->sharedFormulas[$si])) {
            $formula = $this->sharedFormulas[$si];
        }

        return $formula;
    }

    /**
     * @return string
     */
    public function id(): string
    {
        return $this->sheetId;
    }

    /**
     * @return string
     */
    public function name(): string
    {
        return $this->name;
    }

    /**
     * @return string
     */
    public function path(): string
    {
        return $this->pathInZip;
    }

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
     * @param string $state
     *
     * @return $this
     */
    public function setState(string $state): Sheet
    {
        $this->state = $state;

        return $this;
    }

    /**
     * @return string
     */
    public function state(): string
    {
        return $this->state;
    }

    /**
     * @return bool
     */
    public function isVisible(): bool
    {
        return !$this->state || $this->state === 'visible';
    }

    /**
     * @return bool
     */
    public function isHidden(): bool
    {
        return $this->state === 'hidden' || $this->state === 'veryHidden';
    }

    /**
     * @param string|null $file
     *
     * @return Reader
     */
    protected function getReader(?string $file = null): InterfaceXmlReader
    {
        if (empty($this->xmlReader)) {
            if (!$file) {
                $file = $this->zipFilename;
            }
            $this->xmlReader = Excel::createReader($file);
        }

        return $this->xmlReader;
    }

    protected function _readHeader()
    {
        if (!isset($this->dimension['range'])) {
            $this->dimension = [
                'range' => '',
            ];
            $xmlReader = $this->getReader();
            $xmlReader->openZip($this->pathInZip);
            while ($xmlReader->read()) {
                if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'dimension') {
                    $range = (string)$xmlReader->getAttribute('ref');
                    if ($range) {
                        $this->dimension = Helper::rangeArray($range);
                        $this->dimension['range'] = $range;
                    }
                }
                if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'sheetView') {
                    $this->active = (int)$xmlReader->getAttribute('tabSelected');
                }
                if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'col') {
                    if ($xmlReader->hasAttributes) {
                        $colAttributes = [];
                        while ($xmlReader->moveToNextAttribute()) {
                            $colAttributes[$xmlReader->name] = $xmlReader->value;
                        }
                        $this->cols[] = $colAttributes;
                        $xmlReader->moveToElement();
                    }

                }
                if ($xmlReader->name === 'sheetData') {
                    break;
                }
            }
            $xmlReader->close();
        }
    }

    protected function _readBottom()
    {
        if ($this->mergedCells === null) {
            $xmlReader = $this->getReader();
            $xmlReader->openZip($this->pathInZip);
            while ($xmlReader->read()) {
                if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'sheetData') {
                    break;
                }
            }
            $this->mergedCells = [];
            while ($xmlReader->read()) {
                if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'mergeCell') {
                    $ref = (string)$xmlReader->getAttribute('ref');
                    if ($ref) {
                        $arr = Helper::rangeArray($ref);
                        $this->mergedCells[$arr['min_cell']] = $ref;
                    }
                }
            }
            $xmlReader->close();
        }
    }

    /**
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
        $areaRange = $range ?: $this->dimension();
        if ($areaRange && preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            return (int)$matches[5] - (int)$matches[2] + 1;
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
            return Excel::colNum($matches[4]) - Excel::colNum($matches[1]) + 1;
        }

        return 0;
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
     * @return array
     */
    public function getColAttributes(): array
    {
        $result = [];
        if ($this->cols) {
            foreach ($this->cols as $colAttributes) {
                if (isset($colAttributes['min'])) {
                    $col = Helper::colLetter($colAttributes['min']);
                    $result[$col] = $colAttributes;
                }
                else {
                    $result[] = $colAttributes;
                }
            }
        }

        return $result;
    }

    /**
     * @param $dateFormat
     *
     * @return $this
     */
    public function setDateFormat($dateFormat): Sheet
    {
        $this->excel->setDateFormat($dateFormat);

        return $this;
    }

    protected static function _areaRange(string $areaRange): array
    {
        $area = [];
        $area['col_keys'] = [];
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
     * setReadArea('C3:AZ28') - set top left and right bottom of read area
     * setReadArea('C3') - set top left only
     *
     * @param string $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return $this
     */
    public function setReadArea(string $areaRange, ?bool $firstRowKeys = false): Sheet
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
     * setReadArea('C:AZ') - set left and right columns of read area
     * setReadArea('C') - set left column only
     *
     * @param string $columnsRange
     * @param bool|null $firstRowKeys
     *
     * @return $this
     */
    public function setReadAreaColumns(string $columnsRange, ?bool $firstRowKeys = false): Sheet
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
     * Returns values, styles and other info of cells as array
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
        $this->readCallback(static function($row, $col, $val) use (&$columnKeys, &$rowData, &$rowNum) {
            if ($rowNum === -1) {
                $rowNum = $row;
            }
            elseif ($rowNum !== $row) {
                return true;
            }
            $rowData[$col . $row] = $val;

            return null;
        }, $columnKeys, null, $styleIdxInclude);

        return $rowData;
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

        return $this->readRows($columnKeys, $resultMode | Excel::KEYS_SWAP);
    }

    /**
     * Returns values and styles of cells as array ['v' => _value_, 's' => _styles_]
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
     * Returns values and styles of cells as array:
     *      'v' => _value_
     *      's' => _styles_
     *      'f' => _formula_
     *      't' => _type_
     *      'o' => _original_value_
     *
     * @param $styleKey
     *
     * @return array
     */
    public function readCellsWithStyles($styleKey = null): array
    {
        $data = $this->readCells(true);
        foreach ($data as $cell => $cellData) {
            if (isset($cellData['s'])) {
                $style = $this->excel->getCompleteStyleByIdx($cellData['s']);
                if ($styleKey && isset($style[$styleKey])) {
                    $data[$cell]['s'] = [$styleKey => $style[$styleKey]];
                }
                else {
                    $data[$cell]['s'] = $style;
                }
            }
        }

        return $data;
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
                if (isset($this->area['col_keys']) && array_key_exists($col, $this->area['col_keys'])
                    || (!is_array($val) && $val !== null) || isset($val['v']) || isset($val['f']) || isset($val['s'])) {
                    $needBreak = $callback($row, $col, $val);
                    if (!isset($this->area['first_row'])) {
                        $this->area['first_row'] = $row;
                        $this->area['first_col'] = $col;
                    }
                    if ($needBreak) {
                        return;
                    }
                }
            }
        }
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
        elseif ($columnKeys === true) {
            $firstRowKeys = true;
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

        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->pathInZip);

        $rowData = $rowTemplate;
        $rowNum = 0;
        $rowOffset = $colOffset = null;
        $row = -1;
        $rowCnt = -1;

        if ($this->preReadFunc) {
            ($this->preReadFunc)($xmlReader);
        }

        if ($xmlReader->seekOpenTag('sheetData')) {
            while ($xmlReader->read()) {
                if ($rowLimit > 0 && $rowCnt >= $rowLimit) {
                    break;
                }
                if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'sheetData') {
                    break;
                }
                if ($this->readNodeFunc && isset($this->readNodeFunc[$xmlReader->name])) {
                    ($this->readNodeFunc[$xmlReader->name])($xmlReader->expand());
                }

                if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'row') {
                    //$this->countReadRows++;
                    if ($rowNum >= $readArea['row_min'] && $rowNum <= $readArea['row_max']) {
                        $this->readRowNum = $rowNum;
                        if ($rowCnt === 0 && $firstRowKeys) {
                            if (!$columnKeys) {
                                if ($styleIdxInclude) {
                                    $columnKeys = array_combine(array_keys($rowData), array_column($rowData, 'v'));
                                }
                                else {
                                    $columnKeys = $rowData;
                                }
                                $rowTemplate = array_fill_keys(array_keys($columnKeys), null);
                            }
                        }
                        else {
                            if ($resultMode & Excel::RESULT_MODE_ROW) {
                                $rowNode = $xmlReader->expand();
                                $rowAttributes = [];
                                foreach ($rowNode->attributes as $key => $val) {
                                    $rowAttributes[$key] = $val->value;
                                }
                                $rowData = [
                                    '__cells' => $rowData,
                                    '__row' => $rowAttributes,
                                ];
                            }
                            $row = $rowNum - $rowOffset;
                            yield $row => $rowData;
                        }
                        continue;
                    }
                }

                if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                    if ($xmlReader->name === 'row') { // <row ...> - tag row begins
                        $rowNum = (int)$xmlReader->getAttribute('r');

                        if ($rowNum > $readArea['row_max']) {
                            break;
                        }
                        if ($rowNum < $readArea['row_min']) {
                            continue;
                        }
                        $rowData = $rowTemplate;

                        $rowCnt += 1;
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
                        if ($xmlReader->isEmptyElement && ($resultMode & Excel::RESULT_MODE_ROW)) {
                            $rowNode = $xmlReader->expand();
                            $rowAttributes = [];
                            foreach ($rowNode->attributes as $key => $val) {
                                $rowAttributes[$key] = $val->value;
                            }
                            $rowData = [
                                '__cells' => $rowData,
                                '__row' => $rowAttributes,
                            ];
                            $row = $rowNum - $rowOffset;
                            yield $row => $rowData;
                        }
                    } // <row ...> - tag row end

                    elseif ($xmlReader->name === 'c') { // <c ...> - tag cell begins
                        $addr = $xmlReader->getAttribute('r');
                        if ($addr && preg_match('/^([A-Za-z]+)(\d+)$/', $addr, $m)) {
                            //
                            if ($m[2] < $readArea['row_min'] || $m[2] > $readArea['row_max']) {
                                continue;
                            }
                            $colLetter = $m[1];
                            $colNum = Excel::colNum($colLetter);

                            if ($colNum >= $readArea['col_min'] && $colNum <= $readArea['col_max']) {
                                if ($colOffset === null) {
                                    $colOffset = $colNum - 1;
                                    if (is_int($resultMode) && ($resultMode & Excel::KEYS_COL_ZERO_BASED)) {
                                        $colOffset += 1;
                                    }
                                }
                                if ($resultMode) {
                                    if (!($resultMode & (Excel::KEYS_COL_ZERO_BASED | Excel::KEYS_COL_ONE_BASED))) {
                                        $col = $colLetter;
                                    }
                                    else {
                                        $col = $colNum - $colOffset;
                                    }
                                }
                                else {
                                    $col = $colLetter;
                                }
                                $cell = $xmlReader->expand();
                                if (is_array($columnKeys) && isset($columnKeys[$colLetter])) {
                                    $col = $columnKeys[$colLetter];
                                }
                                ///$value = $this->_cellValue($cell, $styleIdx, $formula, $dataType, $originalValue);
                                $value = $this->_cellValue($cell, $additionalData);
                                if ($styleIdxInclude) {
                                    $rowData[$col] = $additionalData;
                                }
                                else {
                                    if (is_string($value) && ($resultMode & Excel::TRIM_STRINGS)) {
                                        $value = trim($value);
                                    }
                                    if (!($value === '' && ($resultMode & Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL))) {
                                        $rowData[$col] = $value;
                                    }
                                }
                            }
                        }
                    } // <c ...> - tag cell end
                }
            }
        }

        if ($this->postReadFunc) {
            ($this->postReadFunc)($xmlReader);
        }

        $xmlReader->close();

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
     * @return int
     */
    public function getReadRowNum(): int
    {
        return $this->readRowNum;
    }

    /**
     * Returns all merged ranges
     *
     * @return array|null
     */
    public function getMergedCells(): ?array
    {
        if ($this->mergedCells === null) {
            $this->_readBottom();
        }

        return $this->mergedCells;
    }

    /**
     * Checks if a cell is merged
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
     * Returns merge range of specified cell
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

    /**
     * @return string|null
     */
    protected function drawingFilename(): ?string
    {
        $findName = str_replace('/worksheets/sheet', '/drawings/drawing', $this->pathInZip);

        return in_array($findName, $this->excel->innerFileList(), true) ? $findName : null;
    }

    /**
     * @param string $cell
     * @param string $fileName
     * @param string|null $imageName
     *
     * @return void
     */
    protected function addImage(string $cell, string $fileName, ?string $imageName = null)
    {
        $this->images[$cell] = [
            'image_name' => $imageName,
            'file_name' => $fileName,
        ];
    }

    /**
     * @param $xmlName
     *
     * @return array
     */
    protected function extractDrawingInfo($xmlName): array
    {
        $drawings = [
            'xml' => $xmlName,
            'rel' => dirname($xmlName) . '/_rels/' . basename($xmlName) . '.rels',
        ];
        $contents = file_get_contents('zip://' . $this->zipFilename . '#' . $xmlName);
        $typeAnchors = [];
        if (preg_match_all('#<xdr:oneCellAnchor[^>]*>(.*)</xdr:oneCellAnchor#siU', $contents, $anchors)) {
            $typeAnchors['one'] = $anchors[1];
        }
        if (preg_match_all('#<xdr:twoCellAnchor[^>]*>(.*)</xdr:twoCellAnchor#siU', $contents, $anchors)) {
            $typeAnchors['two'] = $anchors[1];
        }
        if (preg_match_all('#<xdr:absoluteAnchor>[^>]*>(.*)</xdr:absoluteAnchor>#siU', $contents, $anchors)) {
            $typeAnchors['abs'] = $anchors[1];
        }
        foreach ($typeAnchors as $type => $anchors) {
            foreach ($anchors as $anchorStr) {
                $picture = [];
                if (preg_match('#<xdr:pic>(.*)</xdr:pic>#siU', $anchorStr, $pic)) {
                    if (preg_match('#<a:blip\s(.*)r:embed="(.+)"#siU', $pic[1], $m)) {
                        $picture['rId'] = $m[2];
                    }
                    if ($picture && preg_match('#<xdr:cNvPr(.*)\sname="([^"]*)"/?>#siU', $pic[1], $m)) {
                        $picture['name'] = $m[2];
                    }
                }
                if ($picture) {
                    if (preg_match('#<xdr:from[^>]*>(.*)</xdr:from#siU', $anchorStr, $m)) {
                        if (preg_match('#<xdr:col>(.*)</xdr:col#siU', $m[1], $m1)) {
                            $picture['colIdx'] = (int)$m1[1];
                            $picture['col'] = Excel::colLetter($picture['colIdx'] + 1);
                        }
                        if (preg_match('#<xdr:row>(.*)</xdr:row#siU', $m[1], $m1)) {
                            $picture['rowIdx'] = (int)$m1[1];
                            $picture['row'] = (string)($picture['rowIdx'] + 1);
                        }
                    }
                    if (isset($picture['col'], $picture['row'])) {
                        $picture['cell'] = $picture['col'] . $picture['row'];
                        $drawings['media'][$picture['rId']] = $picture;
                    }
                }
            }
        }

        if (!empty($drawings['media'])) {
            $contents = file_get_contents('zip://' . $this->zipFilename . '#' . $drawings['rel']);
            if (preg_match_all('#<Relationship\s([^>]+)>#siU', $contents, $rel)) {
                foreach ($rel[1] as $str) {
                    if (preg_match('#Id="(\w+)#', $str, $m1) && preg_match('#Target="([^"]+)#', $str, $m2)) {
                        $rId = $m1[1];
                        if (isset($drawings['media'][$rId])) {
                            $drawings['media'][$rId]['target'] = str_replace('../', 'xl/', $m2[1]);
                        }
                    }
                }
            }
        }

        $result = [
            'xml' => $drawings['xml'],
            'rel' => $drawings['rel'],
        ];
        foreach ($drawings['media'] as $media) {
            if (isset($media['target'])) {
                $addr = $media['col'] . $media['row'];
                if (!isset($media['name'])) {
                    $media['name'] = $addr;
                }
                $result['images'][$addr] = $media;
                $result['rows'][$media['row']][] = $addr;
                $this->addImage($addr, basename($media['target']), $media['name']);
            }
        }

        return $result;
    }

    protected function extractRichValueImages()
    {
        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->pathInZip);
        while ($xmlReader->read()) {
            // seek <sheetData>
            if ($xmlReader->name === 'sheetData') {
                break;
            }
        }
        while ($xmlReader->read()) {
            // loop until </sheetData>
            if ($xmlReader->name === 'sheetData' && $xmlReader->nodeType === \XMLReader::END_ELEMENT) {
                break;
            }
            if ($xmlReader->name === 'c' && $xmlReader->nodeType === \XMLReader::ELEMENT) {
                $vm = (string)$xmlReader->getAttribute('vm');
                $cell = (string)$xmlReader->getAttribute('r');
                if ($vm && ($imageFile = $this->excel->metadataImage($vm))) {
                    $this->addImage($cell, basename($imageFile));
                }
            }
        }
        $xmlReader->close();
    }

    /**
     * @return bool
     */
    public function hasDrawings(): bool
    {
        return (bool)$this->drawingFilename();
    }

    /**
     * Count images of the sheet
     *
     * @return int
     */
    public function countImages(): int
    {
        if ($this->countImages === -1) {
            $this->_countDrawingsImages();
            if ($this->excel->hasExtraImages()) {
                $this->extractRichValueImages();
            }
            $this->countImages = count($this->images);
        }

        return $this->countImages;
    }

    /**
     * Count images form drawings of the sheet
     *
     * @return int
     */
    public function _countDrawingsImages(): int
    {
        $result = 0;
        if ($this->hasDrawings()) {
            if (!isset($this->props['drawings'])) {
                if ($xmlName = $this->drawingFilename()) {
                    $this->props['drawings'] = $this->extractDrawingInfo($xmlName);
                }
                else {
                    $this->props['drawings'] = [];
                }
            }
            if (!empty($this->props['drawings']['images'])) {
                $result = count($this->props['drawings']['images']);
            }
        }

        return $result;
    }

    /**
     * @return array
     */
    public function _getDrawingsImageFiles(): array
    {
        $result = [];
        if ($this->_countDrawingsImages()) {
            $result = array_column($this->props['drawings']['images'], 'target');
        }

        return $result;
    }

    /**
     * @return array
     */
    public function getImageList(): array
    {
        if ($this->countImages()) {
            return $this->images;
        }

        return [];
    }

    /**
     * @param $row
     *
     * @return array
     */
    public function getImageListByRow($row): array
    {
        $result = [];
        if ($this->countImages()) {
            if (isset($this->props['drawings']['rows'][$row])) {
                foreach ($this->props['drawings']['rows'][$row] as $addr) {
                    $result[$addr] = [
                        'image_name' => $this->props['drawings']['images'][$addr]['name'],
                        'file_name' => basename($this->props['drawings']['images'][$addr]['target']),
                    ];
                }
            }
        }

        return $result;
    }

    /**
     * Returns TRUE if the cell contains an image
     *
     * @param string $cell
     *
     * @return bool
     */
    public function hasImage(string $cell): bool
    {
        if ($this->countImages()) {
            return isset($this->images[strtoupper($cell)]);
        }

        return false;
    }

    /**
     * Returns full path of an image from the cell (if exists) or null
     *
     * @param string $cell
     *
     * @return string|null
     */
    public function imageEntryFullPath(string $cell): ?string
    {
        if ($this->countImages()) {
            $cell = strtoupper($cell);
            if (isset($this->props['drawings']['images'][$cell])) {

                return 'zip://' . $this->zipFilename . '#' . $this->props['drawings']['images'][$cell]['target'];
            }
        }

        return null;
    }

    /**
     * Returns the MIME type for an image from the cell as determined by using information from the magic.mime file
     * Requires fileinfo extension
     *
     * @param string $cell
     *
     * @return string|null
     */
    public function getImageMimeType(string $cell): ?string
    {
        if (function_exists('mime_content_type') && ($path = $this->imageEntryFullPath($cell))) {
            return mime_content_type($path);
        }

        return null;
    }

    /**
     * Returns the name for an image from the cell as it defines in XLSX
     *
     * @param string $cell
     *
     * @return string|null
     */
    public function getImageName(string $cell): ?string
    {
        if ($this->countImages()) {
            $cell = strtoupper($cell);
            if (isset($this->props['drawings']['images'][$cell])) {

                return $this->props['drawings']['images'][$cell]['name'];
            }
        }

        return null;
    }

    /**
     * Returns an image from the cell as a blob (if exists) or null
     *
     * @param string $cell
     *
     * @return string|null
     */
    public function getImageBlob(string $cell): ?string
    {
        if ($path = $this->imageEntryFullPath($cell)) {
            return file_get_contents($path);
        }

        return null;
    }

    /**
     * Writes an image from the cell to the specified filename
     *
     * @param string $cell
     * @param string|null $filename
     *
     * @return string|null
     */
    public function saveImage(string $cell, ?string $filename = null): ?string
    {
        if ($contents = $this->getImageBlob($cell)) {
            if (!$filename) {
                $filename = basename($this->props['drawings']['images'][strtoupper($cell)]['target']);
            }
            if (file_put_contents($filename, $contents)) {
                return realpath($filename);
            }
        }

        return null;
    }

    /**
     * Writes an image from the cell to the specified directory
     *
     * @param string $cell
     * @param string $dirname
     *
     * @return string|null
     */
    public function saveImageTo(string $cell, string $dirname): ?string
    {
        $filename = basename($this->props['drawings']['images'][strtoupper($cell)]['target']);

        return $this->saveImage($cell, str_replace(['\\', '/'], '', $dirname) . DIRECTORY_SEPARATOR . $filename);
    }

    /**
     * Returns an array of data validation rules found in the sheet
     *
     * @return array<array{
     *   type: string,
     *   sqref: string,
     *   formula1: ?string,
     *   formula2: ?string,
     *  }>
     */
    public function getDataValidations(): array
    {
        if ($this->validations === null) {
            $this->extractDataValidations();
        }

        return $this->validations;
    }

    /** Extracts data validation rules from the sheet */
    public function extractDataValidations(): void
    {
        $validations = [];
        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->pathInZip);

        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                // Standard data validation
                if ($xmlReader->name === 'dataValidation') {
                    $validation = $this->parseDataValidation($xmlReader);
                    if ($validation) {
                        $validations[] = $validation;
                    }
                }

                // Extended data validation
                if ($xmlReader->name === 'x14:dataValidation') {
                    $validation = $this->parseExtendedDataValidation($xmlReader);
                    if ($validation) {
                        $validations[] = $validation;
                    }
                }
            }
        }

        $xmlReader->close();

        $this->validations = $validations;
    }

    /**
     * Parse standard <dataValidation>
     *
     * @param InterfaceXmlReader $xmlReader
     *
     * @return array{
     *    type: string,
     *    sqref: string,
     *    formula1: ?string,
     *    formula2: ?string,
     *  }
     */
    protected function parseDataValidation(InterfaceXmlReader $xmlReader): ?array
    {
        $type = $xmlReader->getAttribute('type');
        $sqref = $xmlReader->getAttribute('sqref');
        $formula1 = null;
        $formula2 = null;

        // Handle child nodes like formula1 and formula2
        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'formula1') {
                $xmlReader->read();
                $formula1 = $xmlReader->value;
            } elseif ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'formula2') {
                $xmlReader->read();
                $formula2 = $xmlReader->value;
            }
            if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'dataValidation') {
                break;
            }
        }

        return [
            'type' => $type,
            'sqref' => $sqref,
            'formula1' => $formula1,
            'formula2' => $formula2
        ];
    }

    /**
     * Parse extended <x14:dataValidation>
     *
     * @param InterfaceXmlReader $xmlReader
     *
     * @return array{
     *    type: string,
     *    sqref: string,
     *    formula1: ?string,
     *    formula2: ?string,
     *  }
     */
    protected function parseExtendedDataValidation(InterfaceXmlReader $xmlReader): array
    {
        $type = $xmlReader->getAttribute('type');
        $sqref = null;
        $formula1 = null;
        $formula2 = null;

        // Parse the attributes within the <x14:dataValidation> tag
        while ($xmlReader->read()) {
            // Parse the sqref (cell range)
            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'xm:sqref') {
                $xmlReader->read();
                $sqref = $xmlReader->value;
            }

            // Capture formula1 and extract inner <xm:f> value
            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'x14:formula1') {
                while ($xmlReader->read()) {
                    if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'xm:f') {
                        $xmlReader->read();
                        $formula1 = $xmlReader->value;
                        break;
                    }
                }
            }

            // Capture formula2 and extract inner <xm:f> value
            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'x14:formula2') {
                while ($xmlReader->read()) {
                    if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'xm:f') {
                        $xmlReader->read();
                        $formula2 = $xmlReader->value;
                        break;
                    }
                }
            }

            // Break when reaching the end of <x14:dataValidation>
            if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'x14:dataValidation') {
                break;
            }
        }

        return [
            'type' => $type,
            'sqref' => $sqref,
            'formula1' => $formula1,
            'formula2' => $formula2
        ];
    }

    /**
     * Returns an array of data validation rules found in the sheet
     *
     * @return array<array{
     *   type: string,
     *   sqref: string,
     *   attributes: array
     * }>
     */
    public function getConditionalFormatting(): array
    {
        if ($this->conditionals === null) {
            $this->extractConditionalFormatting();
        }

        return $this->conditionals;
    }

    /** Extracts conditional formatting rules from the sheet */
    public function extractConditionalFormatting(): void
    {
        $conditionals = [];
        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->pathInZip);

        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'conditionalFormatting') {
                $conditional = $this->parseConditionalFormatting($xmlReader);
                if ($conditional) {
                    $conditionals[] = $conditional;
                }
            }
        }

        $xmlReader->close();

        $this->conditionals = $conditionals;
    }

    /**
     * Parse <conditionalFormatting>
     *
     * @param InterfaceXmlReader $xmlReader
     *
     * @return array{
     *    type: string,
     *    sqref: string,
     *    attributes: []
     *  }
     */
    protected function parseConditionalFormatting(InterfaceXmlReader $xmlReader): ?array
    {
        $sqref = $xmlReader->getAttribute('sqref');
        $attributes = [];

        // Handle child nodes like formula1 and formula2
        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'cfRule') {
                $node = $xmlReader->expand();
                foreach ($node->attributes as $key => $val) {
                    $attributes[$key] = $val->value;
                }
            }
            if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'conditionalFormatting') {
                break;
            }
        }

        return [
            'type' => $attributes['type'] ?? null,
            'sqref' => $sqref,
            'attributes' => $attributes,
        ];
    }

    public function setDefaultRowHeight(float $rowHeight): void
    {
        $this->defaultRowHeight = $rowHeight;
    }

    /**
     * Parses and retrieves column widths and row heights from the sheet XML.
     *
     * @return void
     */
    protected function extractColumnWidthsAndRowHeights(): void
    {
        $this->colWidths = [];
        $this->rowHeights = [];

        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->pathInZip);

        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                // Extract column width
                if ($xmlReader->name === 'col') {
                    $min = (int)$xmlReader->getAttribute('min');
                    $max = (int)$xmlReader->getAttribute('max');
                    $width = (float)$xmlReader->getAttribute('width');

                    for ($i = $min; $i <= $max; $i++) {
                        $this->colWidths[$i] = $width;
                    }
                }
                // Extract row height
                elseif ($xmlReader->name === 'row') {
                    $rowIndex = (int)$xmlReader->getAttribute('r');
                    $height = $xmlReader->getAttribute('ht') ? (float)$xmlReader->getAttribute('ht') : $this->defaultRowHeight;
                    $this->rowHeights[$rowIndex] = $height;
                }
            }
        }

        $xmlReader->close();
    }

    /**
     * Returns column width for a specific column number.
     *
     * @param int $colNumber
     * @return float|null
     */
    public function getColumnWidth(int $colNumber): ?float
    {
        if ($this->colWidths === null) {
            $this->extractColumnWidthsAndRowHeights();
        }
        return $this->colWidths[$colNumber] ?? null;
    }

    /**
     * Returns row height for a specific row number.
     *
     * @param int $rowNumber
     *
     * @return float|null
     */
    public function getRowHeight(int $rowNumber): ?float
    {
        if ($this->rowHeights === null) {
            $this->extractColumnWidthsAndRowHeights();
        }
        return $this->rowHeights[$rowNumber] ?? null;
    }

    /**
     * Parses and retrieves frozen pane info from the sheet XML
     *
     * @return array|null
     */
    public function getFreezePaneInfo(): ?array
    {
        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->pathInZip);

        $freezePane = null;

        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'pane') {
                $xSplit = (int)$xmlReader->getAttribute('xSplit');
                $ySplit = (int)$xmlReader->getAttribute('ySplit');
                $topLeftCell = $xmlReader->getAttribute('topLeftCell');

                $freezePane = [
                    'xSplit' => $xSplit,
                    'ySplit' => $ySplit,
                    'topLeftCell' => $topLeftCell,
                ];
                break;
            }
        }
        $xmlReader->close();

        return $freezePane;
    }

    /**
     * Alias of getFreezePaneInfo()
     *
     * @return array|null
     */
    public function getFreezePaneConfig0(): ?array
    {
        return $this->readCells();
    }

    /**
     * Extracts the tab properties from the sheet XML
     *
     * @return void
     */
    protected function _readTabProperties(): void
    {
        if ($this->tabProperties !== null) {
            return;
        }

        $this->tabProperties = [
            'color' => null,
        ];

        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->pathInZip);

        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'sheetPr') {
                while ($xmlReader->read()) {
                    if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'tabColor') {
                        $this->tabProperties['color'] = [
                            'rgb' => $xmlReader->getAttribute('rgb'),
                            'theme' => $xmlReader->getAttribute('theme'),
                            'tint' => $xmlReader->getAttribute('tint'),
                            'indexed' => $xmlReader->getAttribute('indexed'),
                        ];

                        $this->tabProperties['color'] = array_filter(
                            $this->tabProperties['color'],
                            static fn($value) => $value !== null
                        );
                        break;
                    }
                    if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'sheetPr') {
                        break;
                    }
                }
                break;
            }
        }

        $xmlReader->close();
    }

    /**
     * Returns the tab color info of the sheet
     * Contains any of: rgb, theme, tint, indexed
     *
     * @return array|null
     */
    public function getTabColorInfo(): ?array
    {
        if ($this->tabProperties === null) {
            $this->_readTabProperties();
        }

        return $this->tabProperties['color'] ?? null;
    }

    /**
     * Alias of getTabColorConfig()
     *
     * @return array|null
     */
    public function getTabColorConfiguration(): ?array
    {
        return $this->getTabColorInfo();
    }
}