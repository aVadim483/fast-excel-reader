<?php

namespace avadim\FastExcelReader;

class Sheet
{
    public Excel $excel;

    protected string $zipFilename;

    protected string $sheetId;

    protected string $name;

    protected string $path;

    protected array $area = [];

    protected array $props = [];

    /** @var Reader */
    protected Reader $xmlReader;

    public function __construct($file, $sheetId, $name, $path)
    {
        $this->zipFilename = $file;
        $this->sheetId = $sheetId;
        $this->name = $name;
        $this->path = $path;

        $this->area = [
            'row_min' => 1,
            'col_min' => 1,
            'row_max' => Excel::EXCEL_2007_MAX_ROW,
            'col_max' => Excel::EXCEL_2007_MAX_COL,
            'first_row' => false,
        ];
    }

    /**
     * @param $cell
     *
     * @return mixed
     */
    protected function _cellValue($cell)
    {
        // Determine data type
        $dataType = (string)$cell->getAttribute('t');
        $cellValue = null;
        foreach($cell->childNodes as $node) {
            if ($node->nodeName === 'v') {
                $cellValue = $node->nodeValue;
                break;
            }
        }

        $format = null;
        if ( $dataType === '' || $dataType === 'n' ) { // number
            $styleIdx = (int)$cell->getAttribute('s');
            if ($styleIdx > 0 && ($style = $this->excel->styleByIdx($styleIdx))) {
                $format = $style['format'] ?? null;
                if (isset($style['formatType'])) {
                    $dataType = $style['formatType'];
                }
            }
        }

        $value = '';

        switch ( $dataType ) {
            case 's':
                // Value is a shared string
                if (is_numeric($cellValue) && ($str = $this->excel->sharedString((int)$cellValue)) !== null) {
                    $value = $str;
                }
                break;

            case 'b':
                // Value is boolean
                $value = (bool)$cellValue;
                break;

            case 'inlineStr':
                // Value is rich text inline
                $value = $cell->textContent;
                break;

            case 'e':
                // Value is an error message
                $value = (string)$cellValue;
                break;

            case 'd':
                // Value is a date and non-empty
                if (!empty($cellValue)) {
                    $value = $this->excel->formatDate(Excel::timestamp($cellValue));
                }
                break;

            default:
                // Value is a string
                $value = (string) $cellValue;

                // Check for numeric values
                if (is_numeric($value) && $dataType !== 's') {
                    /** @noinspection TypeUnsafeComparisonInspection */
                    if ($value == (int)$value) {
                        $value = (int)$value;
                    }
                    /** @noinspection TypeUnsafeComparisonInspection */
                    elseif ($value == (float)$value) {
                        $value = (float)$value;
                    }
                }
        }

        return $value;
    }


    public function name(): string
    {
        return $this->name;
    }
    
    
    public function isName($name): bool
    {
        return strcasecmp($this->name, $name) === 0;
    }
    
    protected function getReader($file = null): Reader
    {
        if (empty($this->xmlReader)) {
            if (!$file) {
                $file = $this->zipFilename;
            }
            $this->xmlReader = new Reader($file);
        }

        return $this->xmlReader;
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
        if (preg_match('/^([A-Z]+)(\d+)(:([A-Z]+)(\d+))?$/', $areaRange, $matches)) {
            $this->area['col_min'] = Excel::colNum($matches[1]);
            $this->area['row_min'] = (int)$matches[2];
            if (empty($matches[3])) {
                $this->area['col_max'] = Excel::EXCEL_2007_MAX_COL;
                $this->area['row_max'] = Excel::EXCEL_2007_MAX_ROW;
            }
            else {
                $this->area['col_max'] = Excel::colNum($matches[4]);
                $this->area['row_max'] = (int)$matches[5];
            }
            $this->area['first_row'] = $firstRowKeys;

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
        if (preg_match('/^([A-Z]+)(:([A-Z]+))?$/', $columnsRange, $matches)) {
            $this->area['col_min'] = Excel::colNum($matches[1]);
            if (empty($matches[2])) {
                $this->area['col_max'] = Excel::EXCEL_2007_MAX_COL;
            }
            else {
                $this->area['col_max'] = Excel::colNum($matches[3]);
            }
            $this->area['first_row'] = $firstRowKeys;

            return $this;
        }
        throw new Exception('Wrong address or range "' . $columnsRange . '"');
    }

    /**
     * Returns cell values as a two-dimensional array
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $indexStyle
     *
     * @return array
     */
    public function readRows($columnKeys = [], int $indexStyle = null): array
    {
/*
        if (!is_array($columnKeys)) {
            if (is_int($columnKeys) && $columnKeys > 1 && $indexStyle === null) {
                $firstRowKeys = $columnKeys & Excel::KEYS_FIRST_ROW;
                $indexStyle = $columnKeys;
            }
            else {
                $firstRowKeys = (bool)$columnKeys;
            }
            $columnKeys = [];
        }
        elseif (is_int($indexStyle) && $indexStyle & Excel::KEYS_FIRST_ROW) {
            $firstRowKeys = true;
        }
        else {
            $firstRowKeys = null;
        }

        if ($firstRowKeys === null) {
            $firstRowKeys = !empty($this->area['first_row']);
        }
        if ($columnKeys) {
            $columnKeys = array_combine(array_map('strtoupper', array_keys($columnKeys)), array_values($columnKeys));
        }
        if ($firstRowKeys) {
            $indexStyle = (int)$indexStyle | Excel::KEYS_FIRST_ROW;
        }
*/
        $data = [];
        $this->readCallback(static function($row, $col, $val) use (&$columnKeys, &$data) {
            if (isset($columnKeys[$col])) {
                $data[$row][$columnKeys[$col]] = $val;
            }
            else {
                $data[$row][$col] = $val;
            }
        }, $columnKeys, $indexStyle);

        if ($data && ($indexStyle & Excel::KEYS_SWAP)) {
            $newData = [];
            $rowKeys = array_keys($data);
            foreach (array_keys(reset($data)) as $colKey) {
                $newData[$colKey] = array_combine($rowKeys, array_column($data, $colKey));
            }
            return $newData;
        }

        return $data;
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [col][row]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $indexStyle
     *
     * @return array
     */
    public function readColumns($columnKeys = null, int $indexStyle = null): array
    {
        if (is_int($columnKeys) && $columnKeys > 1 && $indexStyle === null) {
            $indexStyle = $columnKeys | Excel::KEYS_RELATIVE;
            $columnKeys = $columnKeys & Excel::KEYS_FIRST_ROW;
        }
        else {
            $indexStyle = $indexStyle | Excel::KEYS_RELATIVE;
        }

        return $this->readRows($columnKeys, $indexStyle | Excel::KEYS_SWAP);
    }

    /**
     * Returns the values of all cells as array
     *
     * @return array
     */
    public function readCells(): array
    {
        $data = [];
        $this->readCallback(static function($row, $col, $val) use (&$data) {
            $data[$col . $row] = $val;
        });

        return $data;
    }

    /**
     * Reads cell values and passes them to a callback function
     *
     * @param callback $callback Callback function($row, $col, $value)
     * @param array|bool|int|null $columnKeys
     * @param int|null $indexStyle
     */
    public function readCallback(callable $callback, $columnKeys = [], int $indexStyle = null)
    {
        foreach ($this->nextRow($columnKeys, $indexStyle) as $row => $rowData) {
            foreach ($rowData as $col => $val) {
                $needBreak = $callback($row, $col, $val);
                if ($needBreak) {
                    return;
                }
            }
        }
    }

    /**
     * @param array|bool|int|null $columnKeys
     * @param int|null $indexStyle
     *
     * @return \Generator|null
     */
    public function nextRow($columnKeys = [], int $indexStyle = null): ?\Generator
    {
        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->path);
        $readArea = $this->area;

        if (is_array($columnKeys)) {
            $firstRowKeys = is_int($indexStyle) && ($indexStyle & Excel::KEYS_FIRST_ROW);
            $columnKeys = array_combine(array_map('strtoupper', array_keys($columnKeys)), array_values($columnKeys));
        }
        elseif ($columnKeys) {
            $firstRowKeys = true;
            $columnKeys = [];
        }
        else {
            $firstRowKeys = false;
        }

        $rowData = [];
        $rowNum = 0;
        $rowOffset = $colOffset = null;
        $row = -1;
        $rowCnt = -1;
        if ($xmlReader->seekOpenTag('sheetData')) {
            while ($xmlReader->read()) {
                if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'sheetData') {
                    break;
                }
                if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                    if ($xmlReader->name === 'row') {
                        $rowCnt += 1;
                        $rowNum = (int)$xmlReader->getAttribute('r');
                        if ($rowOffset === null) {
                            $rowOffset = $rowNum - ($firstRowKeys ? 2 : 1);
                            if (is_int($indexStyle) && ($indexStyle & Excel::KEYS_ROW_ZERO_BASED)) {
                                $rowOffset -= 1;
                            }
                        }
                        if ($rowCnt > 0) {
                            if ($rowCnt === 1 && $firstRowKeys) {
                                if (!$columnKeys) {
                                    $columnKeys = $rowData;
                                }
                                $rowData = [];
                                continue;
                            }
                            yield $row => $rowData;
                            $rowData = [];
                        }
                    }
                    elseif ($xmlReader->name === 'c') {
                        $addr = $xmlReader->getAttribute('r');
                        if ($addr && preg_match('/^([A-Z]+)(\d+)$/', $addr, $m)) {
                            $colLetter = $m[1];
                            $colNum = Excel::colNum($colLetter);

                            if ($colNum >= $readArea['col_min'] && $colNum <= $readArea['col_max']
                                && $rowNum >= $readArea['row_min'] && $rowNum <= $readArea['row_max']) {
                                if ($colOffset === null) {
                                    $colOffset = $colNum - 1;
                                    if (is_int($indexStyle) && ($indexStyle & Excel::KEYS_COL_ZERO_BASED)) {
                                        $colOffset -= 1;
                                    }
                                }
                                if ($indexStyle) {
                                    $row = $rowNum + $rowOffset;
                                    if (!($indexStyle & (Excel::KEYS_COL_ZERO_BASED | Excel::KEYS_COL_ONE_BASED))) {
                                        $col = $colLetter;
                                    }
                                    else {
                                        $col = $colNum + $colOffset;
                                    }
                                }
                                else {
                                    $row = (string)$rowNum;
                                    $col = $colLetter;
                                }
                                $cell = $xmlReader->expand();
                                if (is_array($columnKeys) && isset($columnKeys[$colLetter])) {
                                    $col = $columnKeys[$colLetter];
                                }
                                $rowData[$col] = $this->_cellValue($cell);
                            }
                        }
                    }
                }
            }
        }
        if ($row > -1) {
            yield $row => $rowData;
        }

        $xmlReader->close();

        return null;
    }

    /**
     * @return string|null
     */
    protected function drawingFilename(): ?string
    {
        $findName = str_replace('/worksheets/sheet', '/drawings/drawing', $this->path);

        return in_array($findName, $this->excel->innerFileList(), true) ? $findName : null;
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
        if (preg_match_all('#<xdr:twoCellAnchor[^>]*>(.*)</xdr:twoCellAnchor#siU', $contents, $anchors)) {
            foreach ($anchors[1] as $twoCellAnchor) {
                $drawing = [];
                if (preg_match('#<xdr:pic>(.*)</xdr:pic>#siU', $twoCellAnchor, $pic)) {
                    if (preg_match('#<a:blip\s(.*)r:embed="(.+)"#siU', $twoCellAnchor, $m)) {
                        $drawing['rId'] = $m[2];
                    }
                    if ($drawing && preg_match('#<xdr:cNvPr(.*)\sname="(.*)">#siU', $pic[1], $m)) {
                        $drawing['name'] = $m[2];
                    }
                }
                if ($drawing) {
                    if (preg_match('#<xdr:from[^>]*>(.*)</xdr:from#siU', $twoCellAnchor, $m)) {
                        if (preg_match('#<xdr:col>(.*)</xdr:col#siU', $m[1], $m1)) {
                            $drawing['colIdx'] = (int)$m1[1];
                            $drawing['col'] = Excel::colLetter($drawing['colIdx'] + 1);
                        }
                        if (preg_match('#<xdr:row>(.*)</xdr:row#siU', $m[1], $m1)) {
                            $drawing['rowIdx'] = (int)$m1[1];
                            $drawing['row'] = (string)($drawing['rowIdx'] + 1);
                        }
                    }
                    $drawings['media'][$drawing['rId']] = $drawing;
                    if (isset($drawing['col'], $drawing['row'])) {
                        $drawing['cell'] = $drawing['col'] . $drawing['row'];
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
            }
        }

        return $result;
    }

    /**
     * @return bool
     */
    public function hasDrawings(): bool
    {
        return (bool)$this->drawingFilename();
    }

    /**
     * @return int
     */
    public function countImages(): int
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
    public function getImageList(): array
    {
        $result = [];
        if ($this->countImages()) {
            foreach ($this->props['drawings']['images'] as $addr => $image) {
                $result[$addr] = [
                    'image_name' => $image['name'],
                    'file_name' => basename($image['target']),
                ];
            }
        }

        return $result;
    }

    /**
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

            return isset($this->props['drawings']['images'][strtoupper($cell)]);
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
}