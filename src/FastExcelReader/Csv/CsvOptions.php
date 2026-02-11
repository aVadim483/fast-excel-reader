<?php

namespace avadim\FastExcelReader\Csv;

/**
 * Class CsvOptions
 *
 * @package avadim\FastExcelReader
 */
class CsvOptions
{
    const STRICT_MODE = 'strict';
    const TOLERANT_MODE = 'tolerant';

    /** @var string|null Column delimiter (null for auto-detect) */
    public ?string $delimiter = null;

    /** @var string The char that encloses the fields */
    public string $enclosure = '"';

    /** @var bool RFC4180 allows double quotes */
    public bool $doubleQuotes = true;

    /** @var string Escape character, usually '\' or '' ('' for no escape) */
    public string $escape = '';

    /** @var bool Trim spaces around values (does not affect spaces inside quotes) */
    public bool $trimFields  = true;

    public ?string $encoding = null;
    public string $mode = self::STRICT_MODE;

    /**
     * @param array $options
     *
     * @return static
     */
    public static function create(array $options = []): CsvOptions
    {
        $instance = new static();
        foreach ($options as $key => $value) {
            if (property_exists($instance, $key)) {
                $instance->$key = $value;
            }
        }

        return $instance;
    }

    /**
     * @param string|null $delimiter
     *
     * @return $this
     */
    public function setDelimiter(?string $delimiter): CsvOptions
    {
        $this->delimiter = $delimiter;

        return $this;
    }

    /**
     * @param string $enclosure
     *
     * @return $this
     */
    public function setEnclosure(string $enclosure): CsvOptions
    {
        $this->enclosure = $enclosure;

        return $this;
    }

    /**
     * @param string $escape
     *
     * @return $this
     */
    public function setEscape(string $escape): CsvOptions
    {
        $this->escape = $escape;

        return $this;
    }

    /**
     * @param string $encoding
     *
     * @return $this
     */
    public function setEncoding(string $encoding): CsvOptions
    {
        $this->encoding = $encoding;

        return $this;
    }

    /**
     * @param bool $enable
     *
     * @return $this
     */
    public function setDoubleQuotes(bool $enable): CsvOptions
    {
        $this->doubleQuotes = $enable;

        return $this;
    }

    /**
     * @param bool $enable
     *
     * @return $this
     */
    public function setTrimFields(bool $enable): CsvOptions
    {
        $this->trimFields = $enable;

        return $this;
    }

    /**
     * @param string $mode
     *
     * @return $this
     */
    public function setMode(string $mode): CsvOptions
    {
        $this->mode = $mode;

        return $this;
    }
}
