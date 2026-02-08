<?php

namespace avadim\FastExcelReader;

/**
 * Class CsvOptions
 *
 * @package avadim\FastExcelReader
 */
class CsvOptions
{
    public string $delimiter = ',';
    public string $enclosure = '"';
    public string $escape = '\\';
    public ?string $encoding = null;

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
     * @param string $delimiter
     *
     * @return $this
     */
    public function setDelimiter(string $delimiter): CsvOptions
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
}
