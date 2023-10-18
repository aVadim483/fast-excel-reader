<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceBookReader;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;
use avadim\FastExcelReader\Interfaces\InterfaceXmlReader;

/**
 * Class Excel
 *
 * @package avadim\FastExcelReader
 */
class Excel implements InterfaceBookReader
{
    public const EXCEL_2007_MAX_ROW = 1048576;
    public const EXCEL_2007_MAX_COL = 16384;

    public const KEYS_ORIGINAL = 0;
    public const KEYS_FIRST_ROW = 1;
    public const KEYS_ROW_ZERO_BASED = 2;
    public const KEYS_COL_ZERO_BASED = 4;
    public const KEYS_ZERO_BASED = 6;
    public const KEYS_ROW_ONE_BASED = 8;
    public const KEYS_COL_ONE_BASED = 16;
    public const KEYS_ONE_BASED = 24;
    public const KEYS_RELATIVE = 32;
    public const KEYS_SWAP = 64;

    // nextRow() returns cells & row attributes
    // ['__cells' => [...], '__row' => [...]]
    public const RESULT_MODE_ROW = 1024;

    protected string $file;

    /** @var Reader */
    protected Reader $xmlReader;

    protected array $fileList = [];

    protected array $relations = [];

    protected array $sharedStrings = [];

    protected array $styles = [];

    /** @var Sheet[] */
    protected array $sheets = [];

    protected int $defaultSheetId;

    protected ?string $dateFormat = null;

    /** @var \Closure|callable|null  */
    protected $dateFormatter = null;

    protected bool $date1904 = false;
    protected string $timezone;

    protected array $builtinFormats = [];

    protected array $names = [];

    protected ?array $themeColors = null;

    /**
     * Excel constructor
     *
     * @param string|null $file
     */
    public function __construct(string $file = null)
    {
        if ($file) {
            $this->file = $file;
            $this->_prepare($file);
        }
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

        if (class_exists('IntlDateFormatter')) {
            $formatter = new \IntlDateFormatter(null, \IntlDateFormatter::SHORT, \IntlDateFormatter::NONE);
            $this->builtinFormats[14]['pattern'] = str_replace('#', 'yy', str_replace(['M', 'y'], ['m', 'yyyy'], str_replace('yy', '#', $formatter->getPattern())));

            $formatter = new \IntlDateFormatter(null, \IntlDateFormatter::NONE, \IntlDateFormatter::SHORT);
            $this->builtinFormats[20]['pattern'] = str_replace('H', 'h', $formatter->getPattern());

            $formatter = new \IntlDateFormatter(null, \IntlDateFormatter::NONE, \IntlDateFormatter::MEDIUM);
            $this->builtinFormats[21]['pattern'] = str_replace('H', 'h', $formatter->getPattern());

            $this->builtinFormats[22]['pattern'] = $this->builtinFormats[14]['pattern'] . ' ' . $this->builtinFormats[20]['pattern'];
        }
        else {
            $t = mktime(3, 4, 5, 2, 1, 1999);

            $this->builtinFormats[14]['pattern'] = str_replace(['1999', '99', '02', '2', '01', '1'], ['yyyy', 'yy', 'mm', 'm', 'dd', 'd'], strftime('%x', $t));
            $this->builtinFormats[22]['pattern'] = $this->builtinFormats[14]['pattern'] . ' h:mm';
        }

        $this->timezone = date_default_timezone_get();
        $this->dateFormatter = function ($value, $format = null) {
            if ($format || $this->dateFormat) {
                return gmdate($format ?: $this->dateFormat, $value);
            }
            return $value;
        };
    }

    /**
     * @param string $file
     */
    protected function _prepare(string $file)
    {
        $this->xmlReader = static::createReader($file);
        $this->fileList = $this->xmlReader->fileList();
        foreach ($this->fileList as $fileName) {
            if (strpos($fileName, 'xl/drawings/drawing') === 0) {
                $this->relations['drawings'][] = $fileName;
            }
            elseif (strpos($fileName, 'xl/media/') === 0) {
                $this->relations['media'][] = $fileName;
            }
        }

        $innerFile = 'xl/_rels/workbook.xml.rels';
        $this->xmlReader->openZip($innerFile);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'Relationship') {
                $type = basename($this->xmlReader->getAttribute('Type'));
                if ($type) {
                    $this->relations[$type][$this->xmlReader->getAttribute('Id')] = 'xl/' . $this->xmlReader->getAttribute('Target');
                }
            }
        }
        $this->xmlReader->close();

        if (isset($this->relations['worksheet'])) {
            $this->_loadSheets();
        }
        if (isset($this->relations['sharedStrings'])) {
            $this->_loadSharedStrings(reset($this->relations['sharedStrings']));
        }
        if (isset($this->relations['theme'])) {
            $this->_loadThemes(reset($this->relations['theme']));
        }
        if (isset($this->relations['styles'])) {
            $this->_loadStyles(reset($this->relations['styles']));
        }

        if ($this->sheets) {
            // set current sheet
            $this->selectFirstSheet();
        }
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadSheets(string $innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/workbook.xml';
        }
        $this->xmlReader->openZip($innerFile);

        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                if ($this->xmlReader->name === 'workbookPr') {
                    $date1904 = (string)$this->xmlReader->getAttribute('date1904');
                    if ($date1904 === '1' || $date1904 === 'true') {
                        $this->date1904 = true;
                    }
                }
                elseif ($this->xmlReader->name === 'sheet') {
                    $rId = $this->xmlReader->getAttribute('r:id');
                    $sheetId = $this->xmlReader->getAttribute('sheetId');
                    $path = $this->relations['worksheet'][$rId];
                    if ($path) {
                        $sheetName = $this->xmlReader->getAttribute('name');
                        $this->sheets[$sheetId] = static::createSheet($sheetName, $sheetId, $this->file, $this->relations['worksheet'][$rId], $this);
                        //$this->sheets[$sheetId]->excel = $this;
                        if ($this->sheets[$sheetId]->isActive()) {
                            $this->defaultSheetId = $sheetId;
                        }
                    }
                }
                elseif ($this->xmlReader->name === 'definedName') {
                    $name = $this->xmlReader->getAttribute('name');
                    $address = $this->xmlReader->readString();
                    $this->names[$name] = $address;
                }
            }
        }
        $this->xmlReader->close();
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadSharedStrings(string $innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/sharedStrings.xml';
        }
        $this->xmlReader->openZip($innerFile);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'si' && $node = $this->xmlReader->expand()) {
                $this->sharedStrings[] = $node->textContent;
            }
        }
        $this->xmlReader->close();
    }

    /**
     * @param string|null $innerFile
     *
     * @return void
     */
    protected function _loadThemes(string $innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/theme/theme1.xml';
        }
        $this->xmlReader->openZip($innerFile);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->localName === 'clrScheme') {
                break;
            }
        }
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::END_ELEMENT && $this->xmlReader->localName === 'clrScheme') {
                break;
            }
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->localName === 'srgbClr') {
                $this->themeColors[] = '#' . $this->xmlReader->getAttribute('val');
            }
            elseif ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->localName === 'sysClr') {
                if ($this->xmlReader->getAttribute('val') === 'windowText') {
                    $this->themeColors[] = '#ffffff';
                }
                elseif ($this->xmlReader->getAttribute('val') === 'window') {
                    $this->themeColors[] = '#202020';
                }
                elseif ($lastClr = $this->xmlReader->getAttribute('lastClr')) {
                    $this->themeColors[] = '#' . $lastClr;
                }
                else {
                    $this->themeColors[] = '';
                }
            }
        }
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadStyles(string $innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/styles.xml';
        }
        $this->xmlReader->openZip($innerFile);
        $styleType = '';
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                if ($this->xmlReader->name === 'cellStyleXfs' || $this->xmlReader->name === 'cellXfs') {
                    $styleType = $this->xmlReader->name;
                    continue;
                }
                if ($this->xmlReader->name === 'numFmt') {
                    $numFmtId = (int)$this->xmlReader->getAttribute('numFmtId');
                    $formatCode = $this->xmlReader->getAttribute('formatCode');
                    $numFmts[$numFmtId] = $formatCode;
                }
                elseif ($this->xmlReader->name === 'xf') {
                    $numFmtId = (int)$this->xmlReader->getAttribute('numFmtId');
                    $formatCode = $numFmts[$numFmtId] ?? '';
                    if ($this->_isDatePattern($numFmtId, $formatCode)) {
                        $this->styles[$styleType][] = ['format' => $formatCode, 'formatType' => 'd'];
                    }
                    elseif ($formatCode) {
                        $this->styles[$styleType][] = ['format' => $formatCode];
                    }
                    else {
                        $this->styles[$styleType][] = null;
                    }
                }
            }
        }
        $this->xmlReader->close();
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
     * @param $root
     * @param $tagName
     *
     * @return void
     */
    protected function _loadStyleNumFmts($root, $tagName)
    {
        foreach ($this->builtinFormats as $key => $val) {
            $this->styles['_'][$tagName][$key] = [
                'format-num-id' => $key,
                'format-pattern' => $val['pattern'],
                'format-category' => $val['category'],
            ];
        }
        if ($root) {
            foreach ($root->childNodes as $child) {
                $numFmtId = $child->getAttribute('numFmtId');
                $formatCode = $child->getAttribute('formatCode');
                if ($numFmtId !== '' && $formatCode !== '') {
                    $node = [
                        'format-num-id' => (int)$numFmtId,
                        'format-pattern' => $formatCode,
                        'format-category' => $this->_isDatePattern($numFmtId, $formatCode) ? 'date' : '',
                    ];
                    $this->styles['_'][$tagName][$node['format-num-id']] = $node;
                }
            }
        }
    }

    /**
     * @param $root
     * @param $tagName
     *
     * @return void
     */
    protected function _loadStyleFonts($root, $tagName)
    {
        foreach ($root->childNodes as $font) {
            $node = [];
            foreach ($font->childNodes as $fontStyle) {
                if ($fontStyle->nodeName === 'b') {
                    $node['font-style-bold'] = 1;
                }
                elseif ($fontStyle->nodeName === 'u') {
                    $node['font-style-underline'] = ($fontStyle->getAttribute('formatCode') === 'double' ? 2 : 1);
                }
                elseif ($fontStyle->nodeName === 'i') {
                    $node['font-style-italic'] = 1;
                }
                elseif ($fontStyle->nodeName === 'strike') {
                    $node['font-style-strike'] = 1;
                }
                elseif ($fontStyle->nodeName === 'color') {
                    $color = $this->_extractColor($fontStyle);
                    if ($color) {
                        $node['font-color'] = $color;
                    }
                }
                elseif (($v = $fontStyle->getAttribute('val')) !== '') {
                    if ($fontStyle->nodeName === 'sz') {
                        $name = 'font-size';
                    }
                    else {
                        $name = 'font-' . $fontStyle->nodeName;
                    }
                    $node[$name] = $v;
                }
            }
            $this->styles['_'][$tagName][] = $node;
        }
    }

    /**
     * @param $root
     * @param $tagName
     *
     * @return void
     */
    protected function _loadStyleFills($root, $tagName)
    {
        foreach ($root->childNodes as $fill) {
            $node = [];
            foreach ($fill->childNodes as $patternFill) {
                if (($v = $patternFill->getAttribute('patternType')) !== '') {
                    $node['fill-pattern'] = $v;
                }
                foreach ($patternFill->childNodes as $child) {
                    if ($child->nodeName === 'fgColor') {
                        $color = $this->_extractColor($child);
                        if ($color) {
                            $node['fill-color'] = $color;
                        }
                    }
                }
            }
            $this->styles['_'][$tagName][] = $node;
        }
    }

    /**
     * @param $node
     *
     * @return string
     */
    protected function _extractColor($node): string
    {
        if ($rgb = $node->getAttribute('rgb')) {
            return '#' . substr($rgb, 2);
        }
        $theme = $node->getAttribute('theme');
        if ($theme !== null && $theme !== '') {
            $color = $this->themeColors[(int)$theme] ?? '';
            if ($color) {
                $tint = $node->getAttribute('tint');
                if (!empty($tint)) {
                    $color = Helper::correctColor($color, $tint);
                }
            }
            return $color;
        }

        return '';
    }

    /**
     * @param $root
     * @param $tagName
     *
     * @return void
     */
    protected function _loadStyleBorders($root, $tagName)
    {
        foreach ($root->childNodes as $border) {
            $node = [];
            foreach ($border->childNodes as $side) {
                if (($v = $side->getAttribute('style')) !== '') {
                    $node['border-' . $side->nodeName . '-style'] = $v;
                }
                else {
                    $node['border-' . $side->nodeName . '-style'] = null;
                }
                foreach ($side->childNodes as $child) {
                    if ($child->nodeName === 'color') {
                        $node['border-' . $side->nodeName . '-color'] = '#' . substr($child->getAttribute('rgb'), 2);
                    }
                }
            }
            $this->styles['_'][$tagName][] = $node;
        }
    }

    /**
     * @param $root
     * @param $tagName
     *
     * @return void
     */
    protected function _loadStyleCellXfs($root, $tagName)
    {
        $attributes = ['numFmtId', 'fontId', 'fillId', 'borderId', 'xfId'];
        foreach ($root->childNodes as $xf) {
            $node = [];
            foreach ($attributes as $attribute) {
                if (($v = $xf->getAttribute($attribute)) !== '') {
                    if (substr($attribute, -2) === 'Id') {
                        $node[$attribute] = (int)$v;
                    }
                    else {
                        $node[$attribute] = $v;
                    }
                }
            }
            foreach ($xf->childNodes as $child) {
                if ($child->nodeName === 'alignment') {
                    if ($v = $child->getAttribute('horizontal')) {
                        $node['format']['format-align-horizontal'] = $v;
                    }
                    if ($v = $child->getAttribute('vertical')) {
                        $node['format']['format-align-vertical'] = $v;
                    }
                    if (($v = $child->getAttribute('wrapText')) && ($v === 'true')) {
                        $node['format']['format-wrap-text'] = 1;
                    }
                }
            }
            $this->styles['_'][$tagName][] = $node;
        }
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadCompleteStyles(string $innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/styles.xml';
        }
        $this->xmlReader->openZip($innerFile);

        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                switch ($this->xmlReader->name) {
                    case 'numFmts':
                        $this->_loadStyleNumFmts($this->xmlReader->expand(), 'numFmts');
                        break;
                    case 'fonts':
                        $this->_loadStyleFonts($this->xmlReader->expand(), 'fonts');
                        break;
                    case 'fills':
                        $this->_loadStyleFills($this->xmlReader->expand(), 'fills');
                        break;
                    case 'borders':
                        $this->_loadStyleBorders($this->xmlReader->expand(), 'borders');
                        break;
                    case 'cellStyleXfs':
                        $this->_loadStyleCellXfs($this->xmlReader->expand(), 'cellStyleXfs');
                        break;
                    case 'cellXfs':
                        $this->_loadStyleCellXfs($this->xmlReader->expand(), 'cellXfs');
                        break;
                    default:
                        //
                }
            }
        }
        if (empty($this->styles['_']['numFmts'])) {
            $this->_loadStyleNumFmts(null, 'numFmts');
        }
        $this->xmlReader->close();
    }

    /**
     * Open XLSX file
     *
     * @param string $file
     *
     * @return Excel
     */
    public static function open(string $file): Excel
    {
        return new self($file);
    }

    /**
     * @param string $file
     * @param array|null $errors
     *
     * @return bool
     */
    public static function validate(string $file, ?array &$errors = []): bool
    {
        $result = true;
        $xmlReader = self::createReader($file, [\XMLReader::VALIDATE => true]);

        $fileList = $xmlReader->fileList();
        \libxml_use_internal_errors(true);
        foreach ($fileList as $innerFile) {
            $ext = pathinfo($innerFile, PATHINFO_EXTENSION);
            if (in_array($ext, ['xml', 'rels', 'vml'])) {
                $zipFile = 'zip://' . $file . '#' . $innerFile;
                $dom = new \DOMDocument;
                $dom->load($zipFile);
                $errors = \libxml_get_errors();
                if ($errors) {
                    $result = false;
                }
            }
        }

        return $result;
    }

    /**
     * @param string $sheetName
     * @param $sheetId
     * @param $file
     * @param $path
     * @param $excel
     *
     * @return Sheet
     */
    public static function createSheet(string $sheetName, $sheetId, $file, $path, $excel): InterfaceSheetReader
    {
        return new Sheet($sheetName, $sheetId, $file, $path, $excel);
    }

    /**
     * @param string $file
     * @param array|null $parserProperties
     *
     * @return Reader
     */
    public static function createReader(string $file, ?array $parserProperties = []): InterfaceXmlReader
    {
        return new Reader($file, $parserProperties);
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
     * @param $excelDateTime
     *
     * @return int
     */
    public function timestamp($excelDateTime): int
    {
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
        else {
            if ($this->timezone !== 'UTC') {
                date_default_timezone_set('UTC');
            }
            $t = strtotime($excelDateTime);
            if ($this->timezone !== 'UTC') {
                date_default_timezone_set($this->timezone);
            }
        }

        return (int)$t;
    }

    /**
     * @param string $dateFormat
     *
     * @return $this
     */
    public function setDateFormat(string $dateFormat): Excel
    {
        $this->dateFormat = $dateFormat;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getDateFormat(): ?string
    {
        return $this->dateFormat;
    }

    /**
     * @param $value
     * @param $format
     *
     * @return false|mixed|string
     */
    public function formatDate($value, $format = null, $styleIdx = null)
    {
        if ($this->dateFormatter) {
            return ($this->dateFormatter)($value, $format, $styleIdx);
        }

        return $value;
    }

    /**
     * Sets custom date formatter
     *
     * @param \Closure|callable|string|bool $formatter
     *
     * @return $this
     */
    public function dateFormatter($formatter): Excel
    {
        if ($formatter === false) {
            $this->dateFormatter = null;
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
     * Returns style array by style Idx
     *
     * @param $styleIdx
     *
     * @return array
     */
    public function styleByIdx($styleIdx): array
    {
        return $this->styles['cellXfs'][$styleIdx] ?? [];
    }

    /**
     * Returns string array by index
     *
     * @param $stringId
     *
     * @return string|null
     */
    public function sharedString($stringId): ?string
    {
        return $this->sharedStrings[$stringId] ?? null;
    }

    /**
     * Returns defined names of workbook
     *
     * @return array
     */
    public function getDefinedNames(): array
    {
        return $this->names;
    }

    /**
     * Returns names array of all sheets
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
     * Returns current or specified sheet
     *
     * @param string|null $name
     *
     * @return Sheet|null
     */
    public function sheet(?string $name = null): ?Sheet
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
     * Returns a sheet by name
     *
     * @param string|null $name
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function getSheet(?string $name = null, ?string $areaRange = null, ?bool $firstRowKeys = false): Sheet
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
     * @return Sheet
     */
    public function getSheetById(int $sheetId, string $areaRange = null, ?bool $firstRowKeys = false): Sheet
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
     * @return Sheet
     */
    public function getFirstSheet(string $areaRange = null, ?bool $firstRowKeys = false): Sheet
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
     * @return Sheet
     */
    public function selectSheet(string $name, string $areaRange = null, ?bool $firstRowKeys = false): Sheet
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
     * @return Sheet
     */
    public function selectSheetById(int $sheetId, string $areaRange = null, ?bool $firstRowKeys = false): Sheet
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
     * @return Sheet
     */
    public function selectFirstSheet(string $areaRange = null, ?bool $firstRowKeys = false): Sheet
    {
        $sheet = $this->getFirstSheet($areaRange, $firstRowKeys);
        $this->defaultSheetId = $sheet->id();

        return $sheet;
    }

    /**
     * @param string $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function setReadArea(string $areaRange, ?bool $firstRowKeys = false): Sheet
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
     * Reads cell values and passes them to a callback function
     *
     * @param callback $callback
     * @param int|null $resultMode
     */
    public function readCallback(callable $callback, int $resultMode = null, ?bool $styleIdxInclude = null)
    {
        $this->sheets[$this->defaultSheetId]->readCallback($callback, $resultMode);
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
    public function readRows($columnKeys = [], int $resultMode = null, ?bool $styleIdxInclude = null): array
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
    public function readRowsWithStyles($columnKeys = [], int $resultMode = null): array
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
    public function readColumns($columnKeys = null, int $resultMode = null): array
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
    public function readColumnsWithStyles($columnKeys = null, int $resultMode = null): array
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

    public function innerFileList(): array
    {
        return $this->fileList;
    }

    /**
     * Returns TRUE if the workbook contains an any draw objects (not images only)
     *
     * @return bool
     */
    public function hasDrawings(): bool
    {
        return !empty($this->relations['drawings']);
    }

    /**
     * Returns TRUE if any sheet contains an image object
     *
     * @return bool
     */
    public function hasImages(): bool
    {
        if ($this->hasDrawings()) {
            foreach ($this->sheets as $sheet) {
                if ($sheet->countImages()) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * @return int
     */
    public function countImages(): int
    {
        $result = 0;
        if ($this->hasDrawings()) {
            foreach ($this->sheets as $sheet) {
                $result += $sheet->countImages();
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
        if ($this->hasDrawings()) {
            foreach ($this->sheets as $sheet) {
                $result[$sheet->name()] = $sheet->getImageList();
            }
        }

        return $result;
    }

    public function readStyles(): array
    {
        if (!isset($this->styles['_'])) {
            $this->styles['_'] = [];
            $this->_loadCompleteStyles();
        }

        return $this->styles['_'];
    }

    /**
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
                $result = array_replace_recursive($result, $styles['cellXfs'][$styleIdx]);
            }

            if (isset($result['xfId']) && isset($styles['cellStyleXfs'][$result['xfId']])) {
                if ($styleIdx === 0 || ($styleIdx > 0 && $result['xfId'])) {
                    $result = array_replace_recursive($result, $styles['cellStyleXfs'][$result['xfId']]);
                }
                unset($result['xfId']);
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
     * @param int $styleIdx
     *
     * @return mixed|string
     */
    public function getFormatPattern(int $styleIdx)
    {
        $style = $this->getCompleteStyleByIdx($styleIdx);

        return $style['format']['format-pattern'] ?? '';
    }

    public function _convertDateFormatPattern($pattern)
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

// EOF