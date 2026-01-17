<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelReader\Interfaces\InterfaceXmlReader;

/**
 * Class Reader
 *
 * @package avadim\FastExcelReader
 */
class Reader extends \XMLReader implements InterfaceXmlReader
{
    protected bool $alterMode = false;

    protected string $xlsxFile;

    protected ?string $innerFile = null;

    protected ?\ZipArchive $zip;

    protected array $xmlParserProperties = [];

    /** @var string[] */
    protected array $tmpFiles = [];

    /** @var string|null */
    protected static string $tempDir = '';


    /**
     * @param string $file
     * @param array|null $parserProperties
     */
    public function __construct(string $file, ?array $parserProperties = [])
    {
        $this->xlsxFile = $file;
        $this->zip = new \ZipArchive();
        if ($parserProperties) {
            $this->xmlParserProperties = $parserProperties;
        }
    }

    public function __destruct()
    {
        $this->close();
    }

    /**
     * @param string|null $tempDir
     */
    public static function setTempDir(?string $tempDir = '')
    {
        if ($tempDir) {
            self::$tempDir = $tempDir;
            if (!is_dir($tempDir)) {
                $res = @mkdir($tempDir, 0755, true);
                if (!$res) {
                    throw new Exception('Cannot create directory "' . $tempDir . '"');
                }
            }
            self::$tempDir = realpath($tempDir);
        }
        else {
            self::$tempDir = '';
        }
    }

    /**
     * @return bool|string
     */
    protected function makeTempFile()
    {
        $name = uniqid('xlsx_reader_', true);
        if (!self::$tempDir) {
            $tempDir = sys_get_temp_dir();
            if (!is_writable($tempDir)) {
                $tempDir = getcwd();
            }
        }
        else {
            $tempDir = self::$tempDir;
        }
        $filename = $tempDir . '/' . $name . '.tmp';
        if (touch($filename, time(), time()) && is_writable($filename)) {
            $filename = realpath($filename);
            $this->tmpFiles[] = $filename;
            return $filename;
        }
        else {
            $error = 'Warning: tempdir ' . $tempDir . ' is not writeable';
            if (!self::$tempDir) {
                $error .= ', use ->setTempDir()';
            }
            throw new Exception($error);
        }
    }

    /**
     * Get list of all entries in the ZIP archive
     *
     * @return array
     */
    public function entryList(): array
    {
        $result = [];

        $zip = new \ZipArchive();
        if (defined('\ZipArchive::RDONLY')) {
            $res = $zip->open($this->xlsxFile, \ZipArchive::RDONLY);
        }
        else {
            $res = $zip->open($this->xlsxFile);
        }
        if ($res === true) {
            for ($i = 0; $i < $zip->numFiles; $i++) {
                $result[] = $zip->getNameIndex($i);
            }
            $zip->close();
        }
        else {
            switch ($res) {
                case \ZipArchive::ER_NOENT:
                    $error = 'No such file';
                    $code = $res;
                    break;
                case \ZipArchive::ER_OPEN:
                    $error = 'Can\'t open file';
                    $code = $res;
                    break;
                case \ZipArchive::ER_READ:
                    $error = '';
                    $code = $res;
                    break;
                case \ZipArchive::ER_NOZIP:
                    $error = 'Not a zip archive';
                    $code = $res;
                    break;
                case \ZipArchive::ER_INCONS:
                    $error = 'Zip archive inconsistent';
                    $code = $res;
                    break;
                case \ZipArchive::ER_MEMORY:
                    $error = 'Malloc failure';
                    $code = $res;
                    break;
                default:
                    $error = 'Unknown error';
                    $code = -1;
            }
            $error = 'Error reading file "' . $this->xlsxFile . '" - ' . $error;
            throw new Exception($error, $code);
        }

        return $result;
    }

    /**
     * Get list of files in the ZIP archive
     *
     * @return array
     */
    public function fileList(): array
    {
        $result = [];
        foreach ($this->entryList() as $entry) {
            if (substr($entry, -1) !== '/') {
                $result[] = $entry;
            }
        }

        return $result;
    }

    /**
     * Open an inner file of the ZIP archive
     *
     * @param string $innerFile
     * @param string|null $encoding
     * @param int|null $options
     *
     * @return bool
     */
    public function openZip(string $innerFile, ?string $encoding = null, ?int $options = null): bool
    {
        if ($options === null) {
            $options = 0;
            if (defined('LIBXML_NONET')) {
                $options = $options | LIBXML_NONET;
            }
            if (defined('LIBXML_COMPACT')) {
                $options = $options | LIBXML_COMPACT;
            }
        }
        $result = (!$this->alterMode) && $this->openXmlWrapper($innerFile, $encoding, $options);
        if (!$result) {
            $result = $this->openXmlStream($innerFile, $encoding, $options);
            $this->alterMode = $result;
        }

        return $result;
    }

    /**
     * @param string $innerFile
     * @param string|null $encoding
     * @param int|null $options
     *
     * @return bool
     */
    protected function openXmlWrapper(string $innerFile, ?string $encoding = null, ?int $options = 0): bool
    {
        $this->innerFile = $innerFile;
        $result = @$this->open('zip://' . $this->xlsxFile . '#' . $innerFile, $encoding, $options);
        if ($result) {
            foreach ($this->xmlParserProperties as $property => $value) {
                $this->setParserProperty($property, $value);
            }
        }

        return (bool)$result;
    }

    /**
     * Opens the INTERNAL XML file from XLSX as XMLReader
     * Example: openXml('xl/workbook.xml')
     *
     * @param string $innerPath
     * @param string|null $encoding
     * @param int|null $options
     *
     * @return bool
     */
    protected function openXmlStream(string $innerPath, ?string $encoding = null, ?int $options = 0): bool
    {
        $this->zip = new \ZipArchive();

        if ($this->zip->open($this->xlsxFile) !== true) {
            throw new Exception('Failed to open archive: ' . $this->xlsxFile);
        }

        $st = $this->zip->getStream($innerPath);
        if ($st === false) {
            throw new Exception("Internal file not found: {$innerPath}");
        }

        $tmp = $this->makeTempFile();
        $out = fopen($tmp, 'wb');
        if (!$out) {
            fclose($st);
            throw new Exception("Failed to create temporary file: {$tmp}");
        }

        stream_copy_to_stream($st, $out);
        fclose($st);
        fclose($out);

        if (!$this->open($tmp, $encoding, $options)) {
            throw new Exception("XMLReader::open() failed to open {$tmp}");
        }

        return true;
    }

    /**
     * @return bool
     */
    #[\ReturnTypeWillChange]
    public function close(): bool
    {
        $result = parent::close();
        if ($result) {
            if ($this->innerFile) {
                $this->innerFile = null;
            }
            foreach ($this->tmpFiles as $tmp) {
                if (is_file($tmp)) {
                    @unlink($tmp);
                }
            }
        }

        return $result;
    }

    /**
     * xl/workbook.xml
     *
     * @return bool
     */
    public function openWorkbook(): bool
    {
        return $this->openZip('xl/workbook.xml');
    }

    /**
     * xl/sharedStrings.xml (the file may be missing)
     *
     * @return bool
     */
    public function openSharedStrings(): bool
    {
        return $this->zip->locateName('xl/sharedStrings.xml') !== false
            && $this->openZip('xl/sharedStrings.xml');
    }

    /**
     * Returns a list of sheets from workbook.xml: [[name, sheetId, rId] ...]
     *
     * @return array
     */
    public function sheetList(): array
    {
        $sheets = [];
        $this->openWorkbook();

        while ($this->read()) {
            if ($this->nodeType === \XMLReader::ELEMENT && $this->name === 'sheet') {
                $sheets[] = [
                    'name' => $this->getAttribute('name'),
                    'sheetId' => $this->getAttribute('sheetId'),
                    'rId' => $this->getAttribute('r:id'),
                ];
            }
        }
        $this->close();

        return $sheets;
    }

    /**
     * Open a sheet by index (0..n-1) or name (string).
     * Automatically reads workbook.xml.rels to map rId -> worksheets/sheetN.xml
     *
     * @param int $index
     *
     * @return bool
     */
    public function openSheetByIndex(int $index): bool
    {
        $sheets = $this->sheetList();
        if (!isset($sheets[$index])) {
            throw new Exception("Sheet with index {$index} not found");
        }
        return $this->openSheetByRelId($sheets[$index]['rId']);
    }

    /**
     * @param string $name
     *
     * @return bool
     */
    public function openSheetByName(string $name): bool
    {
        foreach ($this->sheetList() as $s) {
            if ($s['name'] === $name) {
                return $this->openSheetByRelId($s['rId']);
            }
        }
        throw new Exception("Sheet named '{$name}' not found");
    }

    /**
     * Opens xl/_rels/workbook.xml.rels and finds Target by rId
     *
     * @param string $rId
     *
     * @return bool
     */
    protected function openSheetByRelId(string $rId): bool
    {
        $this->openZip('xl/_rels/workbook.xml.rels');
        $target = null;

        while ($this->read()) {
            if ($this->nodeType === \XMLReader::ELEMENT && $this->name === 'Relationship') {
                if ($this->getAttribute('Id') === $rId) {
                    $target = $this->getAttribute('Target'); // "worksheets/sheet1.xml"
                    break;
                }
            }
        }
        $this->close();

        if ($target === null) {
            throw new Exception("Target not found by rId={$rId} in workbook.xml.rels");
        }

        // относительный путь от xl/
        $inner = 'xl/' . ltrim($target, '/');

        return $this->openZip($inner);
    }

    /**
     * @param string $tagName
     *
     * @return bool
     */
    public function seekOpenTag(string $tagName): bool
    {
        while ($this->read()) {
            if ($this->nodeType === \XMLReader::ELEMENT && $this->name === $tagName) {
                return true;
            }
        }
        return false;
    }

    public function validate()
    {
        $this->setParserProperty(self::VALIDATE, true);
        foreach ($this->fileList() as $file) {
            echo $file, '<br>';
        }
    }
}

// EOF