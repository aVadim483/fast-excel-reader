<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceXmlReader;

/**
 * XLSX worksheet reader
 *
 * Implements the format-specific half of AbstractSheet: streaming the sheet XML
 * node by node, typing cell values, and everything that only exists in the
 * OOXML package - drawings, data validations, conditional formatting, column
 * widths, freeze panes and tab properties.
 */
class Sheet extends AbstractSheet
{
    /** Character masks used by strspn() to split a cell address into column and row */
    private const COL_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
    private const ROW_DIGITS = '0123456789';

    protected string $zipFilename;

    protected string $pathInZip;

    /** @var Reader */
    protected InterfaceXmlReader $xmlReader;

    /** @var mixed */
    protected $preReadFunc = null;

    /** @var mixed */
    protected $postReadFunc = null;

    protected array $readNodeFunc = [];

    protected array $sharedFormulas = [];

    protected int $countImages = -1; // -1 - unknown

    protected array $actualRows = [];

    protected array $actualCols = [];

    protected ?array $cellStat = null;

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

        $this->initReadArea();
    }

    /**
     * Get path to the sheet XML file in ZIP archive
     *
     * @return string
     */
    public function path(): string
    {
        return $this->pathInZip;
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
        $dataType = (string)$cell->getAttribute('t');
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
        }
        elseif ($styleIdx) {
            //$cellValue = '';
            // cell is empty, so $cellValue = null;
        }

        // inline rich text is the only case that needs the text of the whole subtree
        $textContent = ($dataType === 'inlineStr') ? $cell->textContent : null;

        return $this->_makeCellValue($dataType, $styleIdx, $cellValue, $formula, $textContent, $additionalData);
    }

    /**
     * Read the cell the reader stands on without materialising a DOM node
     *
     * XMLReader::expand() copies the node into a DOMDocument and wraps it in a
     * DOMElement, so on a sheet of a million cells it is a million allocate/free
     * pairs spent to read two attributes and the text of <v>. Walking the
     * children with read() yields the same raw parts at about half the cost.
     *
     * On return the reader stands on </c> (or is still on <c/> for an empty
     * cell), so the read() loop of the caller continues exactly as before.
     *
     * @param InterfaceXmlReader $xmlReader
     * @param string $address
     * @param array|null $additionalData
     *
     * @return mixed
     */
    protected function _cellValueFast(InterfaceXmlReader $xmlReader, string $address, ?array &$additionalData = [])
    {
        $dataType = (string)$xmlReader->getAttribute('t');
        $styleIdx = (int)$xmlReader->getAttribute('s');

        $cellValue = $formula = $textContent = null;

        if (!$xmlReader->isEmptyElement) {
            $cellDepth = $xmlReader->depth;
            $inlineStr = ($dataType === 'inlineStr');
            $capture = null;
            $formulaText = $formulaType = $formulaSi = $formulaRef = null;

            while ($xmlReader->read()) {
                $nodeType = $xmlReader->nodeType;

                if ($nodeType === \XMLReader::TEXT || $nodeType === \XMLReader::CDATA
                    || $nodeType === \XMLReader::SIGNIFICANT_WHITESPACE || $nodeType === \XMLReader::WHITESPACE) {
                    if ($capture === 'v') {
                        // libxml may split long text into several nodes, and DOMNode::nodeValue
                        // of <v> is the concatenation of them all - so append, do not assign
                        $cellValue .= $xmlReader->value;
                    }
                    elseif ($capture === 'f') {
                        $formulaText .= $xmlReader->value;
                    }
                    if ($inlineStr) {
                        // reproduces DOMNode::textContent of the whole cell
                        $textContent .= $xmlReader->value;
                    }
                    continue;
                }

                if ($nodeType === \XMLReader::ELEMENT) {
                    $capture = null;
                    // only direct children of <c> carry the value and the formula
                    if ($xmlReader->depth === $cellDepth + 1) {
                        if ($cellValue === null && $xmlReader->name === 'v') {
                            $cellValue = '';
                            $capture = 'v';
                        }
                        elseif ($formulaText === null && $xmlReader->name === 'f') {
                            $formulaText = '';
                            $formulaType = (string)$xmlReader->getAttribute('t');
                            $formulaSi = (string)$xmlReader->getAttribute('si');
                            $formulaRef = (string)$xmlReader->getAttribute('ref');
                            $capture = 'f';
                        }
                    }
                    if ($xmlReader->isEmptyElement) {
                        // <v/> and <f/> have no text node to visit and no END_ELEMENT to close them
                        $capture = null;
                    }
                    continue;
                }

                if ($nodeType === \XMLReader::END_ELEMENT) {
                    if ($xmlReader->depth === $cellDepth) { // </c>
                        break;
                    }
                    $capture = null;
                }
            }

            if ($formulaText !== null) {
                $formula = $this->_makeCellFormula($formulaType, $formulaSi, $formulaRef, $formulaText, $address);
            }
        }

        return $this->_makeCellValue($dataType, $styleIdx, $cellValue, $formula, $textContent, $additionalData);
    }

    /**
     * Cast the raw parts of a cell to the resulting value and fill $additionalData
     *
     * Deliberately knows nothing about XML nodes, so that the DOM branch and the
     * streaming fast path share one set of casting rules.
     *
     * @param string $dataType Value of the "t" attribute
     * @param int $styleIdx Value of the "s" attribute
     * @param string|null $cellValue Text of <v>, NULL when the cell has none
     * @param string|null $formula Formula already resolved by _makeCellFormula()
     * @param string|null $textContent Text of the whole cell, needed for inline strings only
     * @param array|null $additionalData
     *
     * @return mixed
     */
    protected function _makeCellValue(string $dataType, int $styleIdx, ?string $cellValue, ?string $formula, ?string $textContent, ?array &$additionalData = [])
    {
        $attributeT = $dataType;
        if ($cellValue === null) {
            $cellValue = $formula;
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
                $value = (string)$textContent;
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
                elseif ($this->excel->getDateFormatter() === null) {
                    $value = $originalValue;
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
        if ($value && $dataType === 'string') {
            $value = Helper::unescapeString($value);
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
        return $this->_makeCellFormula(
            (string)$node->getAttribute('t'),
            (string)$node->getAttribute('si'),
            (string)$node->getAttribute('ref'),
            $node->nodeValue,
            $address
        );
    }

    /**
     * Resolve a formula from the raw attributes and text of <f>, without a DOM node
     *
     * @param string $type Value of the "t" attribute of <f>
     * @param string $si Value of the "si" attribute of <f>
     * @param string $refAttr Value of the "ref" attribute of <f>
     * @param string|null $formula Text of <f>
     * @param string $address
     *
     * @return string
     */
    protected function _makeCellFormula(string $type, string $si, string $refAttr, ?string $formula, string $address): string
    {
        $shared = ($type === 'shared');
        $formula = (string)$formula;
        if ($formula) {
            if ($formula[0] !== '=') {
                $formula = '=' . $formula;
            }
            if ($shared && $si > '') {
                if ($refAttr && preg_match('/^([a-z]+)\$?(\d+)(:\$?([a-z]+)\$?(\d+))?$/i', $refAttr, $m)) {
                    $ref = ['col_num' => Helper::colNumber($m[1]), 'row_num' => (int)$m[2]];
                }
                else {
                    $ref = ['col_num' => 0, 'row_num' => 0];
                }
                $this->sharedFormulas[$si] = ['ref' => $ref, 'formula' => $formula];
            }
        }
        elseif ($shared && $si > '' && isset($this->sharedFormulas[$si])) {
            $formula = $this->sharedFormulas[$si]['formula'];
            $ref = $this->sharedFormulas[$si]['ref'];
            if (preg_match('/^\$?([a-z]+)\$?(\d+)(:\$?([a-z]+)\$?(\d+))?$/i', $address, $m)) {
                $addressNum = [
                    'col_num' => Helper::colNumber($m[1]),
                    'row_num' => (int)$m[2],
                ];
                $formula = preg_replace_callback('/([A-Z]+)([0-9]+)/', function ($matches) use ($addressNum, $ref) {
                    $colNum = Helper::colNumber($matches[1]);
                    $rowNum = (int)$matches[2];
                    $colOffset = $addressNum['col_num'] - $ref['col_num'];
                    $rowOffset = $addressNum['row_num'] - $ref['row_num'];

                    return Helper::colLetter($colNum + $colOffset) . ($rowNum + $rowOffset);
                }, $formula);
            }
        }

        return $formula;
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

    /**
     * @param string $pathInZip
     *
     * @return InterfaceXmlReader|Reader
     */
    protected function xmlReaderOpenZip(string $pathInZip)
    {
        $xmlReader = $this->getReader();
        $xmlReader->openZip($pathInZip);

        return $xmlReader;
    }

    /**
     * @param $xmlReader
     *
     * @return void
     */
    protected function xmlReaderClose(&$xmlReader)
    {
        $xmlReader->close();
        $xmlReader = null;
    }

    protected function _readHeader()
    {
        if (!isset($this->dimension['range'])) {
            $this->dimension = [
                'range' => '',
            ];
            //$xmlReader = $this->getReader();
            //$xmlReader->openZip($this->pathInZip);
            $xmlReader = $this->xmlReaderOpenZip($this->pathInZip);
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
            //$xmlReader->close();
            $this->xmlReaderClose($xmlReader);;
        }
    }

    protected function _readBottom()
    {
        if ($this->mergedCells === null) {
            //$xmlReader = $this->getReader();
            //$xmlReader->openZip($this->pathInZip);
            $xmlReader = $this->xmlReaderOpenZip($this->pathInZip);
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
            //$xmlReader->close();
            $this->xmlReaderClose($xmlReader);;
        }
    }










    /**
     * Scan sheet data and returns actual number of rows and columns
     *
     * Note: this method performs a full streaming pass over the (decompressed) sheet XML,
     * because the ZIP stream of a deflated inner file is not seekable — the tail cannot be
     * reached without reading through the whole entry. On wide sheets requesting columns
     * ($countColumns = true) is significantly more expensive than rows only, since every
     * cell tag must be scanned. Prefer dimension() when the declared range is enough.
     *
     * @param bool $countColumns
     * @param bool $countRows
     * @param int $blockSize
     *
     * @return array
     */
    public function countActualDimension(bool $countColumns = true, bool $countRows = true, int $blockSize = 4096): array
    {
        $needRows = $countRows && !$this->actualRows;
        $needCols = $countColumns && !$this->actualCols;

        // Nothing to compute (both disabled or already cached) — skip opening the stream
        if (!$needRows && !$needCols) {
            return [
                'rows' => $this->actualRows,
                'cols' => $this->actualCols,
            ];
        }

        // Max expected length of an opening <row ...>/<c ...> tag, used to keep tags whole
        // across block boundaries.
        $overlap = 1024;

        $fp = fopen('zip://' . $this->zipFilename . '#' . $this->pathInZip, 'r');
        $minRow = $maxRow = 0;
        $columns = [];
        $carry = '';    // tail of the previous block, to glue tags split on the boundary
        $rowTail = '';  // sliding window of the stream end, to find the last <row>

        while (!feof($fp)) {
            $chunk = fread($fp, $blockSize);
            if ($chunk === false || $chunk === '') {
                break;
            }
            $scan = $carry . $chunk;

            // The first row is searched only until it is found, then we stop scanning the head
            if ($needRows && $minRow === 0) {
                if (preg_match('/<row\s+[^>]*?\br\s*=\s*"?(\d+)"?/', $scan, $m)) {
                    $minRow = (int)$m[1];
                }
            }

            // Columns require the full pass: collect the set of distinct column letters
            if ($needCols) {
                if (preg_match_all('/<c\s+[^>]*?\br\s*=\s*"?([A-Z]+)\d+"?/', $scan, $mm)) {
                    foreach ($mm[1] as $col) {
                        if (!isset($columns[$col])) {
                            $columns[$col] = \avadim\FastExcelHelper\Helper::colNumber($col);
                        }
                    }
                }
            }

            // The last row is taken from the tail window — no per-block regex over rows
            if ($needRows) {
                $rowTail = substr($rowTail . $chunk, -($blockSize + $overlap));
            }

            $carry = substr($scan, -$overlap);
        }
        fclose($fp);

        if ($needRows && preg_match_all('/<row\s+[^>]*?\br\s*=\s*"?(\d+)"?/', $rowTail, $mm)) {
            $maxRow = (int)end($mm[1]);
            if ($minRow === 0) {
                $minRow = (int)reset($mm[1]);
            }
        }

        if ($needCols) {
            asort($columns);
            if ($columns) {
                $this->actualCols['min'] = array_key_first($columns);
                $this->actualCols['max'] = array_key_last($columns);
                $this->actualCols['count'] = $columns[$this->actualCols['max']] - $columns[$this->actualCols['min']] + 1;
            }
        }
        if ($needRows) {
            $this->actualRows['min'] = $minRow;
            $this->actualRows['max'] = $maxRow;
            $this->actualRows['count'] = $maxRow - $minRow + 1;
        }

        return [
            'rows' => $this->actualRows,
            'cols' => $this->actualCols,
        ];
    }

    /**
     * Single streaming pass over the sheet XML collecting rows, columns and cell counts.
     * Results are cached in $actualRows / $actualCols / $cellStat.
     *
     * @param int $blockSize
     *
     * @return void
     */
    protected function _scanStat(int $blockSize = 4096): void
    {
        // Max expected length of an opening tag, used to glue tags split on block boundaries
        $overlap = 1024;

        $fp = fopen('zip://' . $this->zipFilename . '#' . $this->pathInZip, 'r');
        $minRow = $maxRow = 0;
        $columns = [];
        $cntCells = $cntEmpty = 0;
        $carry = '';    // possibly unfinished tag carried to the next iteration
        $rowTail = '';  // sliding window of the stream end, to find the last <row>

        // Counts distinct data (rows/cols/cells) inside a chunk of complete tags
        $count = function (string $scan) use (&$minRow, &$columns, &$cntCells, &$cntEmpty) {
            if ($scan === '') {
                return;
            }
            if ($minRow === 0 && preg_match('/<row\s+[^>]*?\br\s*=\s*"?(\d+)"?/', $scan, $m)) {
                $minRow = (int)$m[1];
            }
            if (preg_match_all('/<c\s[^>]*?\br\s*=\s*"?([A-Z]+)\d+"?/', $scan, $mm)) {
                foreach ($mm[1] as $col) {
                    if (!isset($columns[$col])) {
                        $columns[$col] = \avadim\FastExcelHelper\Helper::colNumber($col);
                    }
                }
            }
            // all cell tags: <c ...>, <c>, <c/>  (word boundary excludes <col>, <cols>, ...)
            $cntCells += preg_match_all('/<c[\s>\/]/', $scan);
            // empty cells: self-closed <c .../> or immediately closed <c ...></c>
            $cntEmpty += preg_match_all('/<c\b[^>]*\/>/', $scan);
            $cntEmpty += preg_match_all('/<c\b[^>]*[^\/]><\/c>/', $scan);
        };

        while (!feof($fp)) {
            $chunk = fread($fp, $blockSize);
            if ($chunk === false || $chunk === '') {
                break;
            }
            $buf = $carry . $chunk;

            // Cut at the last '<': everything before it consists of complete tags only,
            // the remainder (a possibly unfinished tag) is carried over. Prevents both
            // double counting on the overlap and splitting a tag across blocks.
            $cut = strrpos($buf, '<');
            if ($cut === false) {
                $carry = '';
            }
            else {
                $count(substr($buf, 0, $cut));
                $carry = substr($buf, $cut);
            }

            $rowTail = substr($rowTail . $chunk, -($blockSize + $overlap));
        }
        fclose($fp);

        // Process the final remainder (e.g. a sheet ending with a cell tag)
        $count($carry);

        if (preg_match_all('/<row\s+[^>]*?\br\s*=\s*"?(\d+)"?/', $rowTail, $mm)) {
            $maxRow = (int)end($mm[1]);
            if ($minRow === 0) {
                $minRow = (int)reset($mm[1]);
            }
        }

        asort($columns);
        if ($columns) {
            $this->actualCols = [
                'min' => array_key_first($columns),
                'max' => array_key_last($columns),
                'count' => $columns[array_key_last($columns)] - $columns[array_key_first($columns)] + 1,
            ];
        }
        $this->actualRows = [
            'min' => $minRow,
            'max' => $maxRow,
            'count' => $maxRow - $minRow + 1,
        ];
        $this->cellStat = [
            'total' => $cntCells,
            'filled' => $cntCells - $cntEmpty,
        ];
    }

    /**
     * Returns statistics of the sheet: rows, columns and cell counts
     *
     * [
     *      'rows'  => ['min' => int, 'max' => int, 'count' => int],
     *      'cols'  => ['min' => string, 'max' => string, 'count' => int],
     *      'cells' => ['total' => int, 'filled' => int],
     * ]
     *
     * Note: performs a full streaming pass over the sheet XML (cached); memory is O(blockSize),
     * but time grows with the number of cells — expensive on large/wide sheets.
     *
     * @return array
     */
    public function stat(): array
    {
        if ($this->cellStat === null) {
            $this->_scanStat();
        }

        return [
            'rows' => $this->actualRows,
            'cols' => $this->actualCols,
            'cells' => $this->cellStat,
        ];
    }

    /**
     * Returns the actual number of rows from the sheet data area
     *
     * @return int
     */
    public function countActualRows(): int
    {
        if (!$this->actualRows) {
            $this->countActualDimension(false);
        }

        return $this->actualRows['count'] ?? 0;
    }

    /**
     * Get the first actual row number
     *
     * @return int
     */
    public function minActualRow(): int
    {
        if (!$this->actualRows) {
            $this->countActualDimension(false);
        }

        return $this->actualRows['min'] ?? 0;
    }

    /**
     * Get the last actual row number
     *
     * Note: performs a full streaming pass over the sheet (see countActualDimension()).
     *
     * @return int
     */
    public function maxActualRow(): int
    {
        if (!$this->actualRows) {
            $this->countActualDimension(false);
        }

        return $this->actualRows['max'] ?? 0;
    }

    /**
     * Returns the actual number of columns from the sheet data area
     *
     * Note: scans every cell of the sheet (see countActualDimension()); expensive on wide sheets.
     *
     * @return int
     */
    public function countActualColumns(): int
    {
        if (!$this->actualCols) {
            $this->countActualDimension(true, false);
        }

        return $this->actualCols['count'] ?? 0;
    }

    /**
     * Get the first actual column letter
     *
     * @return string
     */
    public function minActualColumn(): string
    {
        if (!$this->actualCols) {
            $this->countActualDimension(true, false);
        }

        return $this->actualCols['min'] ?? '';
    }

    /**
     * Get the last actual column letter
     *
     * @return string
     */
    public function maxActualColumn(): string
    {
        if (!$this->actualCols) {
            $this->countActualDimension(true, false);
        }

        return $this->actualCols['max'] ?? '';
    }

    /**
     * Get the actual dimension range (e.g. "A1:C10")
     *
     * Note: scans every cell of the sheet (see countActualDimension()); expensive on wide sheets.
     *
     * @return string
     */
    public function actualDimension(): string
    {
        $minCell = $maxCell = '';
        $dim = $this->countActualDimension();
        if (isset($dim['rows']['min'], $dim['cols']['min'])) {
            $minCell = $dim['cols']['min'] . $dim['rows']['min'];
        }
        if (isset($dim['rows']['max'], $dim['cols']['max'])) {
            $maxCell = $dim['cols']['max'] . $dim['rows']['max'];
        }
        if ($minCell && !$maxCell) {
            return $minCell;
        }
        if (!$minCell && $maxCell) {
            return $maxCell;
        }

        return $minCell . ':' . $maxCell;
    }

    /**
     * Get all column attributes
     *
     * @return array
     */
    public function getAllColAttributes(): array
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
     * @deprecated at v2.29
     *
     * @return array
     */
    public function getColAttributes(): array
    {
        return $this->getAllColAttributes();
    }

    /**
     * Read all row attributes to the array
     *
     * @return array
     */
    public function getAllRowAttributes(): array
    {
        static $rows = null;

        if ($rows === null) {
            $rows = [];
            $this->reset();
            foreach ($this->nextRow([], Excel::RESULT_MODE_ROW) as $row => $rowData) {
                $rows[$row] = $rowData['__row'];
            }
        }

        return $rows;
    }

    /**
     * Get row attributes (height, style, etc)
     *
     * @param int $row
     *
     * @return array
     */
    public function getRowAttributes(int $row): array
    {
        $allAttributes = $this->getAllRowAttributes();

        return $allAttributes[$row] ?? [];
    }

    /**
     * Get row style
     *
     * @param int $row
     * @param bool|null $flat
     *
     * @return array
     */
    public function getRowStyle(int $row, ?bool $flat = false): array
    {
        $attributes = $this->getRowAttributes($row);
        if (isset($attributes['s'])) {
            return $this->excel->getCompleteStyleByIdx($attributes['s'], $flat);
        }

        return [];
    }
































    /**
     * Walk the sheet and yield one raw row at a time
     *
     * This is the only part of the reading path that knows about the file format.
     * Everything above it - key modes, read areas, result-mode flags, column
     * renaming - is format independent and lives in nextRow().
     *
     * Yields $rowNum => ['cells' => [colLetter => cellData], 'attrs' => array]
     * where cellData is the descriptor produced by _cellValue():
     * ['v' => value, 's' => styleIdx, 'f' => formula, 't' => type, 'o' => original].
     * Cells outside the column range of $readArea are not reported at all, so
     * their values are never even parsed.
     *
     * @param array $readArea
     * @param int $rowLimit
     * @param bool $rowMode TRUE when row attributes are needed and empty rows must be reported
     *
     * @return \Generator|null
     */
    protected function rawRows(array $readArea, int $rowLimit = 0, bool $rowMode = false): ?\Generator
    {
        $xmlReader = $this->xmlReaderOpenZip($this->pathInZip);

        $rowNum = 0;
        $rowCnt = -1;
        $cells = [];

        // readNodeFunc hands raw DOM nodes to user callbacks, which is what expand() is for;
        // everything else can take the streaming path of _cellValueFast()
        $fastPath = !$this->readNodeFunc;

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
                    if ($rowNum >= $readArea['row_min'] && $rowNum <= $readArea['row_max']) {
                        yield $rowNum => [
                            'cells' => $cells,
                            'attrs' => $rowMode ? $this->_rowAttributes($xmlReader) : [],
                        ];
                        $cells = [];

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
                        $cells = [];
                        $rowCnt += 1;

                        // a self-closing <row/> never reaches the END_ELEMENT branch above
                        if ($xmlReader->isEmptyElement && $rowMode) {
                            yield $rowNum => [
                                'cells' => [],
                                'attrs' => $this->_rowAttributes($xmlReader),
                            ];
                        }
                    } // <row ...> - tag row end

                    elseif ($xmlReader->name === 'c') { // <c ...> - tag cell begins
                        $addr = $xmlReader->getAttribute('r');
                        // splitting "AB12" by hand rather than by preg_match(): this runs once
                        // per cell, and a regex here costs more than the whole address check
                        $addrLen = $addr ? strlen($addr) : 0;
                        $letters = $addrLen ? strspn($addr, self::COL_LETTERS) : 0;
                        if ($letters && $letters < $addrLen && strspn($addr, self::ROW_DIGITS, $letters) === $addrLen - $letters) {
                            $colLetter = substr($addr, 0, $letters);
                            $cellRowNum = (int)substr($addr, $letters);
                            if ($cellRowNum < $readArea['row_min'] || $cellRowNum > $readArea['row_max']) {
                                continue;
                            }
                            $colNum = Excel::colNum($colLetter);

                            if ($colNum >= $readArea['col_min'] && $colNum <= $readArea['col_max']) {
                                if ($fastPath) {
                                    $this->_cellValueFast($xmlReader, $addr, $additionalData);
                                }
                                else {
                                    $this->_cellValue($xmlReader->expand(), $additionalData);
                                }
                                $cells[$colLetter] = $additionalData;
                            }
                        }
                    } // <c ...> - tag cell end
                }
            }
        }

        if ($this->postReadFunc) {
            ($this->postReadFunc)($xmlReader);
        }

        //$xmlReader->close();
        $this->xmlReaderClose($xmlReader);;

        return null;
    }

    /**
     * Collect the attributes of the <row> element the reader is positioned on
     *
     * @param InterfaceXmlReader $xmlReader
     *
     * @return array
     */
    protected function _rowAttributes(InterfaceXmlReader $xmlReader): array
    {
        $rowAttributes = [];
        foreach ($xmlReader->expand()->attributes as $key => $attribute) {
            $rowAttributes[$key] = $attribute->value;
        }

        return $rowAttributes;
    }





    /**
     * Get merged cells. Returns an array [min_cell => range]
     *
     * Note: merge definitions live after <sheetData>, so the first call reads the sheet XML
     * through to the end (result is cached). Lazy — only triggered when merged data is requested.
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
     * Get drawing filename
     *
     * @return string|null
     */
    protected function drawingFilename(): ?string
    {
        $findName = str_replace('/worksheets/sheet', '/drawings/drawing', $this->pathInZip);

        return in_array($findName, $this->excel->innerFileList(), true) ? $findName : null;
    }

    /**
     * Add image to the sheet
     *
     * @param string $cell
     * @param string $fileName
     * @param string|null $imageName
     * @param array|null $meta
     *
     * @return void
     */
    protected function addImage(string $cell, string $fileName, ?string $imageName = null, ?array $meta = [])
    {
        $this->images[$cell] = [
            'image_name' => $imageName,
            'file_name' => $fileName,
            'meta' => $meta,
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

    /**
     * @param int $numImages
     *
     * @return void
     */
    protected function extractRichValueImages(int $numImages)
    {
        //$xmlReader = $this->getReader();
        //$xmlReader->openZip($this->pathInZip);
        $xmlReader = $this->xmlReaderOpenZip($this->pathInZip);
        while ($xmlReader->read()) {
            // seek <sheetData>
            if ($xmlReader->name === 'sheetData') {
                break;
            }
        }
        $count = 0;
        while ($xmlReader->read() && $numImages > $count) {
            // loop until </sheetData>
            if ($xmlReader->name === 'sheetData' && $xmlReader->nodeType === \XMLReader::END_ELEMENT) {
                break;
            }
            if ($xmlReader->name === 'c' && $xmlReader->nodeType === \XMLReader::ELEMENT) {
                $vm = (string)$xmlReader->getAttribute('vm');
                $cell = (string)$xmlReader->getAttribute('r');
                if ($vm && ($imageFile = $this->excel->metadataImage($vm))) {
                    $this->addImage($cell, basename($imageFile), null, ['r' => $cell, 'vm' => $vm]);
                    $count++;
                }
            }
        }
        //$xmlReader->close();
        $this->xmlReaderClose($xmlReader);;
    }

    /**
     * Returns true if the sheet has drawings
     *
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
            if ($cnt = $this->excel->countExtraImages()) {
                $this->extractRichValueImages($cnt);
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
     * Get image list
     *
     * @return array
     */
    public function getImageList(): array
    {
        $result = [];
        if ($this->countImages()) {
            foreach ($this->images as $cell => $image) {
                $result[$cell]['image_name'] = $image['image_name'];
                $result[$cell]['file_name'] = $image['file_name'];
            }
        }

        return $result;
    }

    /**
     * Get image list by row number
     *
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
     * Get full path to the image in the ZIP archive
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
     * Get image MIME type
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
     * Get image name
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
     * Get image content as binary string
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
     * Save image to a file
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
     * Save image to a directory
     *
     * @param string $cell
     * @param string $dirname
     *
     * @return string|null
     */
    public function saveImageTo(string $cell, string $dirname): ?string
    {
        $filename = basename($this->props['drawings']['images'][strtoupper($cell)]['target']);

        return $this->saveImage($cell, str_replace(['\\', '/'], DIRECTORY_SEPARATOR, $dirname) . DIRECTORY_SEPARATOR . $filename);
    }

    /**
     * Get data validation rules
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
        //$xmlReader = $this->getReader();
        //$xmlReader->openZip($this->pathInZip);
        $xmlReader = $this->xmlReaderOpenZip($this->pathInZip);

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

        //$xmlReader->close();
        $this->xmlReaderClose($xmlReader);;

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

        // Check if it's a self-closing tag
        if ($xmlReader->isEmptyElement) {
            return [
                'type' => $type,
                'sqref' => $sqref,
                'formula1' => $formula1,
                'formula2' => $formula2
            ];
        }

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

        // Check if it's a self-closing tag
        if ($xmlReader->isEmptyElement) {
            return [
                'type' => $type,
                'sqref' => $sqref,
                'formula1' => $formula1,
                'formula2' => $formula2
            ];
        }

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
     * Get conditional formatting rules
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
        //$xmlReader = $this->getReader();
        //$xmlReader->openZip($this->pathInZip);
        $xmlReader = $this->xmlReaderOpenZip($this->pathInZip);

        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'conditionalFormatting') {
                $conditional = $this->parseConditionalFormatting($xmlReader);
                if ($conditional) {
                    $conditionals[] = $conditional;
                }
            }
        }

        //$xmlReader->close();
        $this->xmlReaderClose($xmlReader);;

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

    /**
     * Set default row height
     *
     * @param float $rowHeight
     */
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

        //$xmlReader = $this->getReader();
        //$xmlReader->openZip($this->pathInZip);
        $xmlReader = $this->xmlReaderOpenZip($this->pathInZip);

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

        //$xmlReader->close();
        $this->xmlReaderClose($xmlReader);;
    }

    /**
     * Get width of the column
     *
     * @param int|string $colNumber
     *
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
     * Get column attributes (width, style, etc)
     *
     * @param int|string $col
     *
     * @return array|mixed
     */
    public function getColumnAttributes($col)
    {
        $allAttributes = $this->getAllColAttributes();
        if (is_numeric($col)) {
            $col = Helper::colLetter($col);
        }
        else {
            $col = strtoupper($col);
        }

        return $allAttributes[$col] ?? [];
    }

    /**
     * Get column style
     *
     * @param int|string $col
     * @param bool|null $flat
     *
     * @return array
     */
    public function getColumnStyle($col, ?bool $flat = false): array
    {
        $attributes = $this->getColumnAttributes($col);
        if (isset($attributes['style'])) {
            return $this->excel->getCompleteStyleByIdx($attributes['style'], $flat);
        }

        return [];
    }

    /**
     * Get height of the row
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
     * Get freeze pane info
     *
     * @return array|null
     */
    public function getFreezePaneInfo(): ?array
    {
        //$xmlReader = $this->getReader();
        //$xmlReader->openZip($this->pathInZip);
        $xmlReader = $this->xmlReaderOpenZip($this->pathInZip);

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
        //$xmlReader->close();
        $this->xmlReaderClose($xmlReader);;

        return $freezePane;
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

        //$xmlReader = $this->getReader();
        //$xmlReader->openZip($this->pathInZip);
        $xmlReader = $this->xmlReaderOpenZip($this->pathInZip);

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

        //$xmlReader->close();
        $this->xmlReaderClose($xmlReader);
    }

    /**
     * Get the tab color info of the sheet
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
     * Get tab color configuration. Alias of getTabColorConfig()
     *
     * @return array|null
     */
    public function getTabColorConfiguration(): ?array
    {
        return $this->getTabColorInfo();
    }
}