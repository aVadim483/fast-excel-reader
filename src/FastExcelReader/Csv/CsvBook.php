<?php

namespace avadim\FastExcelReader\Csv;

use avadim\FastExcelReader\AbstractBook;

/**
 * CSV workbook reader
 *
 * Presents a CSV file through the same book/sheet contract as XLSX and XLS: a
 * workbook with exactly one sheet. It carries no styles, no shared strings and
 * no temp files, so the format-specific hooks are all no-ops - the CsvReader
 * engine does the parsing, and the shared AbstractSheet does the reading.
 *
 * @see \avadim\FastExcelReader\Excel::open() dispatches here for non-XLS/non-ZIP files
 */
class CsvBook extends AbstractBook
{
    protected CsvReader $reader;

    /** @var CsvOptions|array|null */
    protected $csvOptions;

    /**
     * @param string|null $file
     * @param CsvOptions|array|null $options
     */
    public function __construct(?string $file = null, $options = [])
    {
        $this->csvOptions = $options;

        parent::__construct($file);
    }

    /**
     * @param string $file
     *
     * @return void
     */
    protected function _prepare(string $file): void
    {
        $this->reader = new CsvReader($file, $this->csvOptions);
        $sheet = new CsvSheet('CSV', '1', $this->reader, $file, $this);
        $this->sheets = [1 => $sheet];
        $this->defaultSheetId = 1;
    }

    /**
     * CSV has no styles
     *
     * @return void
     */
    protected function _loadCompleteStyles()
    {
        $this->styles['_'] = [
            'numFmts' => [],
            'fonts' => [],
            'fills' => [],
            'borders' => [],
            'cellStyleXfs' => [],
            'cellXfs' => [],
        ];
    }

    /**
     * CSV uses no temporary files
     *
     * @param string $tempDir
     *
     * @return void
     */
    public static function setTempDir($tempDir)
    {
    }

    /**
     * The CSV parsing engine, for low level access
     *
     * @return CsvReader
     */
    public function getReader(): CsvReader
    {
        return $this->reader;
    }
}
