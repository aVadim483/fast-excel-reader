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

    protected array $options = [
        'delimiter' => null,
        'enclosure' => '"',
        'doubleQuotes' => true,
        'escape' => '',
        'trimFields' => true,
        'encoding' => null,
        'mode' => self::STRICT_MODE,
        'stream_filter' => null,
    ];

    /**
     * CsvOptions constructor
     *
     * @param array $options
     */
    public function __construct(array $options = [])
    {
        foreach ($options as $key => $value) {
            $this->__set($key, $value);
        }
    }

    /**
     * Create CsvOptions instance
     *
     * @param array $options
     *
     * @return CsvOptions
     */
    public static function create(array $options = []): CsvOptions
    {
        return new self($options);
    }

    /**
     * Magic setter for options
     *
     * @param string $name
     * @param mixed $value
     *
     * @return void
     */
    public function __set($name, $value)
    {
        if ($name === 'quote') {
            $name = 'enclosure';
        }
        if ($name === 'double_quotes') {
            $name = 'doubleQuotes';
        }
        if ($name === 'trim_fields') {
            $name = 'trimFields';
        }
        if ($name === 'streamFilter') {
            $name = 'stream_filter';
        }
        if (array_key_exists($name, $this->options)) {
            $this->options[$name] = $value;
        }
    }

    /**
     * Magic getter for options
     *
     * @param string $name
     *
     * @return mixed|null
     */
    public function __get($name)
    {
        if ($name === 'quote') {
            $name = 'enclosure';
        }
        if ($name === 'double_quotes') {
            $name = 'doubleQuotes';
        }
        if ($name === 'trim_fields') {
            $name = 'trimFields';
        }
        if ($name === 'streamFilter') {
            $name = 'stream_filter';
        }
        return $this->options[$name] ?? null;
    }

    /**
     * Magic isset for options
     *
     * @param string $name
     *
     * @return bool
     */
    public function __isset($name)
    {
        if ($name === 'quote') {
            $name = 'enclosure';
        }
        if ($name === 'double_quotes') {
            $name = 'doubleQuotes';
        }
        if ($name === 'trim_fields') {
            $name = 'trimFields';
        }
        if ($name === 'streamFilter') {
            $name = 'stream_filter';
        }
        return isset($this->options[$name]);
    }

    /**
     * Set delimiter character
     *
     * @param string|null $delimiter
     *
     * @return $this
     */
    public function setDelimiter(?string $delimiter): CsvOptions
    {
        $this->options['delimiter'] = $delimiter;

        return $this;
    }

    /**
     * Set enclosure character
     *
     * @param string $enclosure
     *
     * @return $this
     */
    public function setEnclosure(string $enclosure): CsvOptions
    {
        $this->options['enclosure'] = $enclosure;

        return $this;
    }

    /**
     * Set escape character
     *
     * @param string $escape
     *
     * @return $this
     */
    public function setEscape(string $escape): CsvOptions
    {
        $this->options['escape'] = $escape;

        return $this;
    }

    /**
     * Set input file encoding (null = auto)
     *
     * @param string $encoding
     *
     * @return $this
     */
    public function setEncoding(string $encoding): CsvOptions
    {
        $this->options['encoding'] = $encoding;

        return $this;
    }

    /**
     * Set whether to handle double quotes
     *
     * @param bool $enable
     *
     * @return $this
     */
    public function setDoubleQuotes(bool $enable): CsvOptions
    {
        $this->options['doubleQuotes'] = $enable;

        return $this;
    }

    /**
     * Set whether to trim fields
     *
     * @param bool $enable
     *
     * @return $this
     */
    public function setTrimFields(bool $enable): CsvOptions
    {
        $this->options['trimFields'] = $enable;

        return $this;
    }

    /**
     * Set parsing mode (strict or tolerant)
     *
     * @param string $mode
     *
     * @return $this
     */
    public function setMode(string $mode): CsvOptions
    {
        $this->options['mode'] = $mode;

        return $this;
    }

    /**
     * Set stream filter
     *
     * @param string|null $filter
     *
     * @return $this
     */
    public function setStreamFilter(?string $filter): CsvOptions
    {
        $this->options['stream_filter'] = $filter;

        return $this;
    }

    /**
     * Return all options as array
     *
     * @return array
     */
    public function toArray(): array
    {
        return $this->options;
    }
}
