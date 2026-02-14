<?php

namespace avadim\FastExcelReader\Csv;

use avadim\FastExcelReader\Excel;

/**
 * Class CsvOptions
 *
 * @package avadim\FastExcelReader
 */
class CsvOptions
{
    public const KEYS_ORIGINAL = Excel::KEYS_ORIGINAL;
    public const KEYS_FIRST_ROW = Excel::KEYS_FIRST_ROW;
    public const KEYS_ROW_ZERO_BASED = Excel::KEYS_ROW_ZERO_BASED;
    public const KEYS_COL_ZERO_BASED = Excel::KEYS_COL_ZERO_BASED;
    public const KEYS_ZERO_BASED = Excel::KEYS_ZERO_BASED;
    public const KEYS_ROW_ONE_BASED = Excel::KEYS_ROW_ONE_BASED;
    public const KEYS_COL_ONE_BASED = Excel::KEYS_COL_ONE_BASED;
    public const KEYS_ONE_BASED = Excel::KEYS_ONE_BASED;

    // Column keys like in Excel: A, B, C, ...
    public const KEYS_COL_EXCEL = 128;

    const STRICT_MODE = 'strict';
    const TOLERANT_MODE = 'tolerant';

    protected array $options = [
        'mode' => self::STRICT_MODE,
        'encoding' => null,
        'delimiter' => null,
        'enclosure' => '"',
        'double_quotes' => true,
        'escape' => '',
        'trim_fields' => true,
        'skip_empty_lines' => true,
        'stream_filter' => null,
        'comment_prefix' => null,
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

    protected static function camelToSnake(string $value): string
    {
        $value = preg_replace('/(?<!^)[A-Z]/', '_$0', $value);

        return strtolower($value);
    }

    protected static function snakeToCamel(string $value): string
    {
        return lcfirst(str_replace(' ', '', ucwords(str_replace('_', ' ', $value))));
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
        $name = self::camelToSnake($name);
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
        $name = self::camelToSnake($name);
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
        $name = self::camelToSnake($name);
        return isset($this->options[$name]);
    }

    /**
     * Set column delimiter character (null for auto-detect)
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
     * Set enclosure character of fields
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
     * Set escape character, usually '\' or '' ('' or null for no escape)
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
        $this->options['double_quotes'] = $enable;

        return $this;
    }

    /**
     * Set whether to trim fields (does not affect spaces inside quotes)
     *
     * @param bool $enable
     *
     * @return $this
     */
    public function setTrimFields(bool $enable): CsvOptions
    {
        $this->options['trim_fields'] = $enable;

        return $this;
    }

    /**
     * Set whether to skip empty lines
     *
     * @param bool $enable
     *
     * @return $this
     */
    public function setSkipEmptyLines(bool $enable): CsvOptions
    {
        $this->options['skip_empty_lines'] = $enable;

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
     * Set comment prefix
     *
     * @param string|null $value
     *
     * @return $this
     */
    public function setCommentPrefix(?string $value): CsvOptions
    {
        $this->options['comment_prefix'] = $value;

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
