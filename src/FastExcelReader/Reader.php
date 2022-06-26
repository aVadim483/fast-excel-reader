<?php

namespace avadim\FastExcelReader;

/**
 * Class Reader
 *
 * @package avadim\FastExcelReader
 */
class Reader extends \XMLReader
{
    protected $zipFile;
    protected $innerFile;

    public function __construct($file)
    {
        $this->zipFile = $file;
    }

    public function __destruct()
    {
        $this->close();
    }

    /**
     * @param string $innerFile
     * @param string|null $encoding
     * @param int|null $options
     *
     * @return bool
     */
    public function openZip($innerFile, $encoding = null, $options = 0)
    {
        $this->innerFile = $innerFile;

        return $this->open('zip://' . $this->zipFile . '#' . $innerFile, $encoding, $options);
    }

    /**
     * @return bool
     */
    public function close()
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
    public function seekOpenTag($tagName)
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