<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Csv\CsvOptions;
use avadim\FastExcelReader\Csv\CsvReader;
use avadim\FastExcelReader\Xls\XlsBook;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;
use avadim\FastExcelReader\Interfaces\InterfaceXmlReader;

/**
 * XLSX workbook reader, and the entry point of the library
 *
 * Implements the format-specific half of AbstractBook: reading the OOXML
 * package - workbook.xml, shared strings, themes, styles and metadata images -
 * and creating XLSX sheet readers.
 *
 * @package avadim\FastExcelReader
 */
class Excel extends AbstractBook
{
    /** @var array Standard Excel indexed colors */
    public const INDEXED_COLORS = [
        0  => 'FF000000', // Black
        1  => 'FFFFFFFF', // White
        2  => 'FFFF0000', // Red
        3  => 'FF00FF00', // Green
        4  => 'FF0000FF', // Blue
        5  => 'FFFFFF00', // Yellow
        6  => 'FFFF00FF', // Magenta
        7  => 'FF00FFFF', // Cyan

        8  => 'FF000000', // Black
        9  => 'FFFFFFFF', // White
        10 => 'FFFF0000', // Red
        11 => 'FF00FF00', // Green
        12 => 'FF0000FF', // Blue
        13 => 'FFFFFF00', // Yellow
        14 => 'FFFF00FF', // Magenta
        15 => 'FF00FFFF', // Cyan

        16 => 'FF800000', // Dark Red
        17 => 'FF008000', // Dark Green
        18 => 'FF000080', // Dark Blue
        19 => 'FF808000', // Olive
        20 => 'FF800080', // Purple
        21 => 'FF008080', // Teal
        22 => 'FFC0C0C0', // Silver
        23 => 'FF808080', // Gray

        24 => 'FF9999FF', // Light Blue
        25 => 'FF993366', // Plum
        26 => 'FFFFFFCC', // Light Yellow
        27 => 'FFCCFFFF', // Light Cyan
        28 => 'FF660066', // Dark Purple
        29 => 'FFFF8080', // Coral
        30 => 'FF0066CC', // Ocean Blue
        31 => 'FFCCCCFF', // Ice Blue

        32 => 'FF000080', // Navy
        33 => 'FFFF00FF', // Pink
        34 => 'FFFFFF00', // Yellow
        35 => 'FF00FFFF', // Cyan
        36 => 'FF800080', // Purple
        37 => 'FF800000', // Brown
        38 => 'FF008080', // Teal
        39 => 'FF0000FF', // Blue

        40 => 'FF00CCFF', // Light Blue
        41 => 'FFCCFFFF', // Aqua
        42 => 'FFCCFFCC', // Light Green
        43 => 'FFFFFF99', // Light Yellow
        44 => 'FF99CCFF', // Sky Blue
        45 => 'FFFF99CC', // Rose
        46 => 'FFCC99FF', // Lavender
        47 => 'FFFFCC99', // Tan

        48 => 'FF3366FF', // Bright Blue
        49 => 'FF33CCCC', // Turquoise
        50 => 'FF99CC00', // Lime
        51 => 'FFFFCC00', // Gold
        52 => 'FFFF9900', // Orange
        53 => 'FFFF6600', // Orange Red
        54 => 'FF666699', // Blue Gray
        55 => 'FF969696', // Gray 40%

        56 => 'FF003366', // Dark Teal
        57 => 'FF339966', // Sea Green
        58 => 'FF003300', // Dark Green
        59 => 'FF333300', // Dark Olive
        60 => 'FF993300', // Brown
        61 => 'FF993366', // Burgundy
        62 => 'FF333399', // Indigo
        63 => 'FF333333', // Gray 80%
    ];


    /** @var Reader */
    protected Reader $xmlReader;

    protected array $fileList = [];

    protected array $relations = [];

    protected array $valueMetadataImages = [];

    protected ?array $themeColors = null;

    protected int $countImages = -1; // -1 - unknown



    /**
     * Set directory for temporary files
     *
     * @param string $tempDir
     */
    public static function setTempDir($tempDir)
    {
        Reader::setTempDir($tempDir);
    }


    /**
     * @param string $file
     */
    protected function _prepare(string $file): void
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
                    $this->relations[$type][$this->xmlReader->getAttribute('Id')] = 'xl/' . ltrim($this->xmlReader->getAttribute('Target'), '/xl');
                }
            }
        }
        $this->xmlReader->close();

        if (isset($this->relations['worksheet'])) {
            $this->_loadSheets();
        }

        if (isset($this->relations['sharedStrings'])) {
            $innerFile = $this->checkInnerFile(reset($this->relations['sharedStrings']));
            if ($innerFile) {
                $this->_loadSharedStrings($innerFile);
            }
        }

        if (isset($this->relations['theme'])) {
            $innerFile = $this->checkInnerFile(reset($this->relations['theme']));
            if ($innerFile) {
                $this->_loadThemes($innerFile);
            }
        }

        if (isset($this->relations['styles'])) {
            $innerFile = $this->checkInnerFile(reset($this->relations['styles']));
            if ($innerFile) {
                $this->_loadStyles($innerFile);
            }
        }

        if (isset($this->relations['sheetMetadata'], $this->relations['richValueRel'])) {
            $metadataFile = $this->checkInnerFile(reset($this->relations['sheetMetadata']));
            $richValueRelFile = $this->checkInnerFile(reset($this->relations['richValueRel']));
            $this->_loadMetadataImages($metadataFile, $richValueRelFile);
        }

        if ($this->sheets) {
            // set current sheet
            $this->selectFirstSheet();
        }
    }

    /**
     * @param string $innerFile
     *
     * @return null|string
     */
    protected function checkInnerFile(string $innerFile): ?string
    {
        foreach ($this->fileList as $filename) {
            if (strcasecmp($innerFile, $filename) === 0) {
                return $filename;
            }
        }
        return null;
    }

    protected function _loadSheets(): void
    {
        $innerFile = $this->checkInnerFile('xl/workbook.xml');
        $this->xmlReader->openZip($innerFile);

        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                $xmlReaderName = $this->xmlReader->name;
                if ($xmlReaderName === 'workbookPr') {
                    $date1904 = (string)$this->xmlReader->getAttribute('date1904');
                    if ($date1904 === '1' || $date1904 === 'true') {
                        $this->date1904 = true;
                    }
                }
                elseif ($xmlReaderName === 'sheet' || $xmlReaderName === 'x:sheet') {
                    $rId = $this->xmlReader->getAttribute('r:id');
                    $sheetId = $this->xmlReader->getAttribute('sheetId');
                    $path = $this->relations['worksheet'][$rId] ?? null;
                    // ignoring non-existent sheets
                    if ($path) {
                        $sheetName = $this->xmlReader->getAttribute('name');
                        $this->sheets[$sheetId] = static::createSheet($sheetName, $sheetId, $this->file, $this->relations['worksheet'][$rId], $this);
                        //$this->sheets[$sheetId]->excel = $this;
                        if ($this->sheets[$sheetId]->isActive()) {
                            $this->defaultSheetId = $sheetId;
                        }
                        if ($state = $this->xmlReader->getAttribute('state')) {
                            $this->sheets[$sheetId]->setState($state);
                        }
                    }
                }
                elseif ($xmlReaderName === 'definedName') {
                    $name = $this->xmlReader->getAttribute('name');
                    $address = $this->xmlReader->readString();
                    $this->names[$name] = $address;
                }
            }
        }
        $this->xmlReader->close();
    }

    /**
     * @param string $innerFile
     */
    protected function _loadSharedStrings(string $innerFile)
    {
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
    protected function _loadThemes(?string $innerFile = null)
    {
        $innerFile = $this->checkInnerFile($innerFile ?: 'xl/theme/theme1.xml');
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
    protected function _loadStyles(?string $innerFile = null)
    {
        $innerFile = $this->checkInnerFile($innerFile ?: 'xl/styles.xml');
        $this->xmlReader->openZip($innerFile);
        $styleType = '';
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                $nodeName = $this->xmlReader->name;
                if ($nodeName === 'cellStyleXfs' || $nodeName === 'cellXfs') {
                    $styleType = $nodeName;
                    continue;
                }
                if ($nodeName === 'numFmt') {
                    $numFmtId = (int)$this->xmlReader->getAttribute('numFmtId');
                    $formatCode = $this->xmlReader->getAttribute('formatCode');
                    $numFmts[$numFmtId] = $formatCode;
                }
                elseif ($nodeName === 'xf') {
                    $numFmtId = (int)$this->xmlReader->getAttribute('numFmtId');
                    $formatCode = $numFmts[$numFmtId] ?? '';
                    if ($this->_isDatePattern($numFmtId, $formatCode)) {
                        $this->styles[$styleType][] = ['format' => $formatCode, 'formatType' => 'd'];
                    }
                    elseif ($formatCode) {
                        if ($this->_isNumberPattern($numFmtId, $formatCode)) {
                            $this->styles[$styleType][] = ['format' => $formatCode, 'formatType' => 'n'];
                        }
                        else {
                            $this->styles[$styleType][] = ['format' => $formatCode];
                        }
                    }
                    elseif ($numFmtId > 0 && isset($this->builtinFormats[$numFmtId]['category'])) {
                        $this->styles[$styleType][] = ['formatType' => $this->builtinFormats[$numFmtId]['category']];
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
     * @param string|null $metadataFile
     */
    protected function _loadMetadataImages(string $metadataFile, string $richValueRelFile)
    {
        $this->xmlReader->openZip($metadataFile);
        $metadataTypesCount = 0;
        $metadataTypes = [];
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->name === 'metadataType') {
                if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                    $metadataTypesCount++;
                    if ((string)$this->xmlReader->getAttribute('name') === 'XLRICHVALUE') {
                        // we need only <metadataType name="XLRICHVALUE" ...>
                        $metadataTypes[$metadataTypesCount] = 'XLRICHVALUE';
                    }
                }
                else {
                    break;
                }
            }
        }
        $futureMetadata = [];
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->name === 'futureMetadata') {
                if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && (string)$this->xmlReader->getAttribute('name') === 'XLRICHVALUE') {
                    while ($this->xmlReader->read()) {
                        if ($this->xmlReader->name === 'xlrd:rvb') {
                            $futureMetadata[] = (int)$this->xmlReader->getAttribute('i');
                        }
                        elseif ($this->xmlReader->name === 'futureMetadata' && $this->xmlReader->nodeType === \XMLReader::END_ELEMENT) {
                            break 2;
                        }
                    }
                }
                elseif ($this->xmlReader->nodeType === \XMLReader::END_ELEMENT) {
                    break;
                }
            }
        }

        while ($this->xmlReader->read()) {
            if ($this->xmlReader->name === 'rc') {
                $type = (int)$this->xmlReader->getAttribute('t');
                $value = (int)$this->xmlReader->getAttribute('v');
                if (isset($metadataTypes[$type])) { // metadataType name="XLRICHVALUE"
                    if (isset($futureMetadata[$value])) {
                        $this->valueMetadataImages[] = ['i' => $futureMetadata[$value]];
                    }
                }
            }
        }
        $this->xmlReader->close();

        $this->xmlReader->openZip($richValueRelFile);
        $count = 0;
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->name === 'rel' && ($rId = $this->xmlReader->getAttribute('r:id'))) {
                $this->valueMetadataImages[$count++]['r_id'] = $rId;
            }
        }
        $this->xmlReader->close();

        $images = [];
        $xmlRels = 'xl/richData/_rels/richValueRel.xml.rels';
        $this->xmlReader->openZip($xmlRels);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->name === 'Relationship' && $this->xmlReader->nodeType === \XMLReader::ELEMENT && ($Id = (string)$this->xmlReader->getAttribute('Id'))) {
                if (substr((string)$this->xmlReader->getAttribute('Type'), -6) === '/image') {
                    $images[$Id] = (string)$this->xmlReader->getAttribute('Target');
                }
            }
        }
        $this->xmlReader->close();

        foreach ($this->valueMetadataImages as $index => $metadataImage) {
            $rId = $this->valueMetadataImages[$index]['r_id'];
            if (isset($images[$rId])) {
                $this->valueMetadataImages[$index]['file_name'] = str_replace('../media/', 'xl/media/', $images[$rId]);
            }
        }
    }

    /**
     * Get image file name from metadata by index
     *
     * @param int $vmIndex
     *
     * @return string|null
     */
    public function metadataImage(int $vmIndex): ?string
    {
        return $this->valueMetadataImages[$vmIndex - 1]['file_name'] ?? null;
    }



    /**
     * Element children of a node, without the text nodes between the tags
     *
     * styles.xml is often pretty-printed, and then childNodes also holds the
     * whitespace between the tags. Those are not styles: they answer no
     * getAttribute(), and counting them would shift the position that every
     * fontId/fillId/borderId/xfId reference relies on.
     *
     * @param \DOMNode $node
     *
     * @return \DOMElement[]
     */
    protected static function _elementChildren($node): array
    {
        $elements = [];
        foreach ($node->childNodes as $child) {
            if ($child->nodeType === XML_ELEMENT_NODE) {
                $elements[] = $child;
            }
        }

        return $elements;
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
            foreach (self::_elementChildren($root) as $child) {
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
        foreach (self::_elementChildren($root) as $font) {
            $node = [];
            foreach (self::_elementChildren($font) as $fontStyle) {
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
        foreach (self::_elementChildren($root) as $fill) {
            $node = [];
            foreach (self::_elementChildren($fill) as $patternFill) {
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
                $tint = (float)$node->getAttribute('tint');
/*
                if (!empty($tint)) {
                    $color0 = Helper::correctColor($color, $tint);
                }
*/
                if ($tint !== 0.0) {
                    if ($tint === 1.0) {
                        $color = '#FFFFFF';
                    }
                    elseif ($tint === -1.0) {
                        $color = '#000000';
                    }
                    else {
                        $r = hexdec(substr($color, 1, 2));
                        $g = hexdec(substr($color, 3, 2));
                        $b = hexdec(substr($color, 5, 2));
                        if ($tint > 0) {
                            $r = round($r + (255 - $r) * $tint);
                            $g = round($g + (255 - $g) * $tint);
                            $b = round($b + (255 - $b) * $tint);
                        }
                        else {
                            $r = round($r * (1 + $tint));
                            $g = round($g * (1 + $tint));
                            $b = round($b * (1 + $tint));
                        }
                        $color = '#'
                            . strtoupper(str_pad(dechex($r), 2, '0', STR_PAD_LEFT))
                            . strtoupper(str_pad(dechex($g), 2, '0', STR_PAD_LEFT))
                            . strtoupper(str_pad(dechex($b), 2, '0', STR_PAD_LEFT));
                    }
                }
            }
            return $color;
        }

        if ($indexed = $node->getAttribute('indexed')) {
            if ($indexed === '64') {
                return '#000000';
            }
            if (isset(self::INDEXED_COLORS[$indexed])) {
                return '#' . substr(self::INDEXED_COLORS[$indexed], 2);
            }
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
        foreach (self::_elementChildren($root) as $border) {
            $node = [];
            foreach (self::_elementChildren($border) as $side) {
                if (($v = $side->getAttribute('style')) !== '') {
                    $node['border-' . $side->nodeName . '-style'] = $v;
                }
                else {
                    $node['border-' . $side->nodeName . '-style'] = null;
                }
                foreach ($side->childNodes as $child) {
                    if ($child->nodeName === 'color') {
                        $node['border-' . $side->nodeName . '-color'] = $this->_extractColor($child);
                        /*
                        if ($attr = $child->getAttribute('rgb')) {
                            $node['border-' . $side->nodeName . '-color'] = '#' . substr($attr, 2);
                        }
                        elseif ($attr = $child->getAttribute('indexed')) {
                            $node['border-' . $side->nodeName . '-color'] = '#' . substr($attr, 2);
                        }
                        else {
                            $node['border-' . $side->nodeName . '-color'] = null;
                        }
                        */
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
        foreach (self::_elementChildren($root) as $xf) {
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
    protected function _loadCompleteStyles(?string $innerFile = null)
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
     * Open a spreadsheet, choosing the reader by the file signature
     *
     * A ZIP container is XLSX, the OLE2 magic number is a legacy XLS workbook. The file extension is not consulted, because it is often wrong on files arriving from other systems.
     *
     * @param string $file
     *
     * @return AbstractBook
     */
    public static function open(string $file): AbstractBook
    {
        if (self::isXls($file)) {
            return XlsBook::open($file);
        }

        return new self($file);
    }

    /**
     * Open an XLS (Excel 97-2003, BIFF8) file
     *
     * @param string $file
     *
     * @return XlsBook
     */
    public static function openXls(string $file): XlsBook
    {
        return XlsBook::open($file);
    }

    /**
     * TRUE if the file starts with the OLE2 compound file signature
     *
     * @param string $file
     *
     * @return bool
     */
    public static function isXls(string $file): bool
    {
        if (!is_readable($file) || is_dir($file)) {
            return false;
        }
        $handle = fopen($file, 'rb');
        if (!$handle) {
            return false;
        }
        $signature = (string)fread($handle, 8);
        fclose($handle);

        return $signature === "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
    }

    /**
     * Open CSV file
     *
     * @param string $file
     * @param CsvOptions|array|null $options
     *
     * @return CsvReader
     */
    public static function openCsv(string $file, $options = []): CsvReader
    {
        return new CsvReader($file, $options);
    }

    /**
     * Validate XLSX file
     *
     * @param string $file
     * @param array|null $errors
     *
     * @return bool
     */
    public static function validate(string $file, ?array &$errors = []): bool
    {
        $result = true;
        $xmlReader = self::createReader($file, [\XMLReader::VALIDATE => true]);

        if (extension_loaded('dom') && extension_loaded('libxml') && function_exists('libxml_use_internal_errors')) {
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
        }

        return $result;
    }

    /**
     * Create sheet object
     *
     * @param string $sheetName
     * @param int|string $sheetId
     * @param string $file
     * @param string $path
     * @param Excel $excel
     *
     * @return InterfaceSheetReader
     */
    public static function createSheet(string $sheetName, $sheetId, $file, $path, $excel): InterfaceSheetReader
    {
        return new Sheet($sheetName, $sheetId, $file, $path, $excel);
    }

    /**
     * Create XML reader object
     *
     * @param string $file
     * @param array|null $parserProperties
     *
     * @return InterfaceXmlReader
     */
    public static function createReader(string $file, ?array $parserProperties = []): InterfaceXmlReader
    {
        return new Reader($file, $parserProperties);
    }




































    /**
     * Get list of inner files in XLSX
     *
     * @return array
     */
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
     * Get list of media image files in the workbook
     *
     * @return array
     */
    public function mediaImageFiles(): array
    {
        $result = [];
        if (!empty($this->relations['media'])) {
            foreach ($this->relations['media'] as $mediaFile) {
                $extension = strtolower(pathinfo($mediaFile, PATHINFO_EXTENSION));
                if (in_array($extension, ['jpg', 'jpeg', 'png', 'bmp', 'ico', 'webp', 'tif', 'tiff', 'gif'])) {
                    $result[] = basename($mediaFile);
                }
            }
        }

        return $result;
    }

    /**
     * Returns the total count of images in the workbook
     *
     * @return int
     */
    public function countImages(): int
    {
        if ($this->countImages === -1) {
            $this->countImages = 0;
            if ($this->hasDrawings() || $this->mediaImageFiles()) {
                foreach ($this->sheets as $sheet) {
                    $this->countImages += $sheet->countImages();
                }
            }
        }

        return $this->countImages;
    }

    /**
     * Get the list of images from the workbook
     *
     * @return array
     */
    public function getImageList(): array
    {
        $result = [];
        if ($this->countImages()) {
            foreach ($this->sheets as $sheet) {
                $result[$sheet->name()] = $sheet->getImageList();
            }
        }

        return $result;
    }

    /**
     * Count "extra" images (images that are in the media folder but not in the drawings)
     *
     * @return int
     */
    public function countExtraImages(): int
    {
        $drawingImageFiles = [];
        if ($this->hasDrawings()) {
            foreach ($this->sheets as $sheet) {
                $imageFiles = $sheet->_getDrawingsImageFiles();
                if ($imageFiles) {
                    $drawingImageFiles += $imageFiles;
                }
            }
        }
        $imageFiles = $this->mediaImageFiles();

        return (count($imageFiles) - count($drawingImageFiles));
    }

    /**
     * Returns TRUE if there are any "extra" images
     *
     * @return bool
     */
    public function hasExtraImages(): bool
    {
        return $this->countExtraImages() > 0;
    }





}

// EOF