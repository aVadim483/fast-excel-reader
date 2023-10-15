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
    protected string $zipFile;

    protected ?string $innerFile = null;

    protected array $xmlParserProperties = [];

    public function __construct($file, ?array $parserProperties = [])
    {
        $this->zipFile = $file;
        if ($parserProperties) {
            $this->xmlParserProperties = $parserProperties;
        }
    }

    public function __destruct()
    {
        $this->close();
    }

    public function entryList(): array
    {
        $result = [];

        $zip = new \ZipArchive();
        if (defined('\ZipArchive::RDONLY')) {
            $res = $zip->open($this->zipFile, \ZipArchive::RDONLY);
        }
        else {
            $res = $zip->open($this->zipFile);
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
            $error = 'Error reading file "' . $this->zipFile . '" - ' . $error;
            throw new Exception($error, $code);
        }

        return $result;
    }

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
     * @param string $innerFile
     * @param string|null $encoding
     * @param int|null $options
     *
     * @return bool
     */
    public function openZip(string $innerFile, string $encoding = null, ?int $options = 0): bool
    {
        $this->innerFile = $innerFile;

        $result = $this->open('zip://' . $this->zipFile . '#' . $innerFile, $encoding, $options);
        foreach ($this->xmlParserProperties as $property => $value) {
            $this->setParserProperty($property, $value);
        }

        return $result;
    }

    /**
     * @return bool
     */
    public function close(): bool
    {
        if ($this->innerFile) {
            $this->innerFile = null;
            return parent::close();
        }
        return true;
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