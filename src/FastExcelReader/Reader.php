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

    protected ?string $innerFile;

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
        if ($zip->open($this->zipFile)) {
            for ($i = 0; $i < $zip->numFiles; $i++) {
                $result[] = $zip->getNameIndex($i);
            }
            $zip->close();
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