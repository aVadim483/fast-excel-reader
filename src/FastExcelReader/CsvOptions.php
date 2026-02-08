<?php

namespace avadim\FastExcelReader;

/**
 * Class CsvOptions
 *
 * @package avadim\FastExcelReader
 */
class CsvOptions
{
    const STRICT_MODE = 'strict';
    const TOLERANT_MODE = 'tolerant';
    public ?string $delimiter = null;
    public string $quote = '"';
    public string $escape = '\\';
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
     * @param string $quote
     *
     * @return $this
     */
    public function setQuote(string $quote): CsvOptions
    {
        $this->quote = $quote;

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
