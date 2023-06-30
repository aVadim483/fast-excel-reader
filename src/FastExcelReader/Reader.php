<?php

namespace avadim\FastExcelReader;

/**
 * Class Reader
 *
 * @package avadim\FastExcelReader
 */
class Reader extends \XMLReader
{
    protected string $zipFile;

    protected ?string $innerFile = null;

    public function __construct($file)
    {
        $this->zipFile = $file;
    }

    public function __destruct()
    {
        $this->close();
    }

    public function fileList(): array
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

        return $this->open('zip://' . $this->zipFile . '#' . $innerFile, $encoding, $options);
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
}

// EOF