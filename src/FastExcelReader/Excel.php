<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Internal\PackagePath;
use avadim\FastExcelReader\Internal\StreamIO;
use avadim\FastExcelReader\Csv\CsvBook;
use avadim\FastExcelReader\Csv\CsvOptions;
use avadim\FastExcelReader\Csv\CsvReader;
use avadim\FastExcelReader\Xls\XlsBook;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;
use avadim\FastExcelReader\Interfaces\InterfaceXmlReader;

use avadim\FastExcelReader\Xlsx\XlsxBook;
use avadim\FastExcelReader\Ods\OdsBook;
use avadim\FastExcelReader\Internal\SpreadsheetPackage;
/** Format-independent spreadsheet opening facade. */
class Excel
{
    /** TRUE for the supported XML workbook family, including macro books and templates. */
    public static function isXlsx(string $file): bool
    {
        return SpreadsheetPackage::isXlsx($file);
    }

    public static function isOds(string $file): bool
    {
        return SpreadsheetPackage::isOds($file);
    }

    public static function openOds(string $file): OdsBook
    {
        return new OdsBook($file);
    }
    public const EXCEL_2007_MAX_ROW = AbstractBook::EXCEL_2007_MAX_ROW;
    public const EXCEL_2007_MAX_COL = AbstractBook::EXCEL_2007_MAX_COL;
    public const KEYS_ORIGINAL = AbstractBook::KEYS_ORIGINAL;
    public const KEYS_FIRST_ROW = AbstractBook::KEYS_FIRST_ROW;
    public const KEYS_ROW_ZERO_BASED = AbstractBook::KEYS_ROW_ZERO_BASED;
    public const KEYS_COL_ZERO_BASED = AbstractBook::KEYS_COL_ZERO_BASED;
    public const KEYS_ZERO_BASED = AbstractBook::KEYS_ZERO_BASED;
    public const KEYS_ROW_ONE_BASED = AbstractBook::KEYS_ROW_ONE_BASED;
    public const KEYS_COL_ONE_BASED = AbstractBook::KEYS_COL_ONE_BASED;
    public const KEYS_ONE_BASED = AbstractBook::KEYS_ONE_BASED;
    public const KEYS_RELATIVE = AbstractBook::KEYS_RELATIVE;
    public const KEYS_SWAP = AbstractBook::KEYS_SWAP;
    public const RESULT_MODE_ROW = AbstractBook::RESULT_MODE_ROW;
    public const TRIM_STRINGS = AbstractBook::TRIM_STRINGS;
    public const TREAT_EMPTY_STRING_AS_EMPTY_CELL = AbstractBook::TREAT_EMPTY_STRING_AS_EMPTY_CELL;
    public const INDEXED_COLORS = XlsxBook::INDEXED_COLORS;

    private function __construct()
    {
    }

    public static function colNum(string $column): int
    {
        return AbstractBook::colNum($column);
    }

    public static function colLetter(int $column): string
    {
        return AbstractBook::colLetter($column);
    }

    public static function setTempDir($directory)
    {
        Reader::setTempDir($directory);
    }
    /**
     * Open a spreadsheet, choosing the reader by the file signature
     *
     * The OLE2 magic number is a legacy XLS workbook, a supported ZIP container is an XML workbook or ODS, and anything else is treated as delimited text (CSV). The file extension is not consulted, because it is often wrong on files arriving from other systems. Pass $options['format'] = 'csv' to force the CSV reader, and any CsvOptions keys (delimiter, enclosure, encoding, ...) to configure it.
     *
     * @param string $file
     * @param CsvOptions|array|null $options
     *
     * @return AbstractBook
     */
    public static function open(string $file, $options = []): AbstractBook
    {
        // An empty file is not a valid spreadsheet or CSV. Report it plainly
        // here, instead of silently opening a zero-row CSV further down (open()
        // dispatches by signature, and an empty file matches neither XLS nor
        // XLSX, so it would otherwise fall through to the CSV backend).
        if (is_file($file) && filesize($file) === 0) {
            throw new Exception("File $file is empty");
        }
        if (is_array($options) && isset($options['format']) && strtolower((string)$options['format']) === 'csv') {
            return new CsvBook($file, $options);
        }
        if (self::isXls($file)) {
            return XlsBook::open($file);
        }
        if (self::isZip($file)) {
            if (self::isOds($file)) {
                return new OdsBook($file);
            }
            SpreadsheetPackage::requireXlsx($file);
            return new XlsxBook($file);
        }

        // Neither OLE2 (XLS) nor ZIP (XLSX): treat as delimited text
        return new CsvBook($file, $options);
    }

    /**
     * Open a spreadsheet held in a string, choosing the reader by its signature
     *
     * The content is written to a temporary file and then opened exactly like open() - format is detected from the bytes, not from any file name, so an XLSX/XLS/CSV payload each read back the same as its on-disk counterpart. The temporary file is owned by the returned workbook and removed on close(), with shutdown cleanup as a fallback. Handy for content coming from a database blob, an HTTP response body or an S3/Flysystem read.
     *
     * @param string $content Raw bytes of the workbook
     * @param CsvOptions|array|null $options Same options as open()
     *
     * @return AbstractBook
     */
    public static function openString(string $content, $options = []): AbstractBook
    {
        if ($content === '') {
            throw new Exception('Cannot open an empty string as a spreadsheet');
        }
        $tempFile = Reader::tempFilename();
        try {
            StreamIO::writeString($tempFile, $content);
        }
        catch (\Throwable $e) {
            @unlink($tempFile);
            throw $e;
        }

        return self::openTempFile($tempFile, $options);
    }

    /**
     * Open a spreadsheet from an open stream resource
     *
     * The stream is copied (from its current position, without seeking, so non-rewindable streams such as HTTP wrappers work) into a temporary file and then opened like open(). This is the entry point for URLs (fopen('https://...')), php://memory and Flysystem/S3 read streams. The caller keeps ownership of the stream; it is not closed here. The temporary file is owned by the returned workbook and removed on close(), with shutdown cleanup as a fallback.
     *
     * @param resource $stream An open, readable stream resource
     * @param CsvOptions|array|null $options Same options as open()
     *
     * @return AbstractBook
     */
    public static function openStream($stream, $options = []): AbstractBook
    {
        if (!is_resource($stream) || get_resource_type($stream) !== 'stream') {
            throw new Exception('openStream() expects an open stream resource');
        }
        $tempFile = Reader::tempFilename();
        try {
            if (StreamIO::copyToFile($stream, $tempFile) === 0) {
                throw new Exception('The stream produced no data');
            }
        }
        catch (\Throwable $e) {
            @unlink($tempFile);
            throw $e;
        }

        return self::openTempFile($tempFile, $options);
    }

    /**
     * Open a temporary file produced by openString()/openStream() and transfer
     * ownership to the workbook
     *
     * @param string $tempFile
     * @param CsvOptions|array|null $options
     *
     * @return AbstractBook
     */
    protected static function openTempFile(string $tempFile, $options): AbstractBook
    {
        // Nothing owns this whole-workbook temporary file (the readers reopen it
        // per inner entry), so remove it when the script ends. A guard keeps this
        // safe if the file was already deleted.
        register_shutdown_function(static function () use ($tempFile) {
            if (is_file($tempFile)) {
                @unlink($tempFile);
            }
        });

        try {
            $book = self::open($tempFile, $options);
            $book->ownTemporaryFile($tempFile);
            return $book;
        }
        catch (\Throwable $e) {
            @unlink($tempFile);
            throw $e;
        }
    }

    /**
     * Open an XLSX workbook without format detection
     *
     * @param string $file
     * @return XlsxBook
     */
    public static function openXlsx(string $file): XlsxBook
    {
        SpreadsheetPackage::requireXlsx($file);
        return new XlsxBook($file);
    }

    /**
     * Open CSV as a single-sheet workbook without format detection
     *
     * Unlike open(), this factory also accepts an empty CSV file.
     *
     * @param string $file
     * @param CsvOptions|array|null $options
     * @return CsvBook
     */
    public static function openCsvBook(string $file, $options = []): CsvBook
    {
        return new CsvBook($file, $options);
    }

    /**
     * Open the low-level CSV reader without a workbook wrapper
     *
     * Use this factory for low-level CSV access. openCsv() and openCsvBook() return a workbook.
     *
     * @param string $file
     * @param CsvOptions|array|null $options
     * @return CsvReader
     */
    public static function openCsvReader(string $file, $options = []): CsvReader
    {
        return new CsvReader($file, $options);
    }

    /**
     * Open an XLS (Excel 97-2003, BIFF8) file
     *
     * @param string $file
     *
     * @return XlsBook
     */
    public static function openXls(string $file): XlsBook
    {
        return XlsBook::open($file);
    }

    /**
     * TRUE if the file starts with the OLE2 compound file signature
     *
     * @param string $file
     *
     * @return bool
     */
    public static function isXls(string $file): bool
    {
        if (!is_readable($file) || is_dir($file)) {
            return false;
        }
        $handle = fopen($file, 'rb');
        if (!$handle) {
            return false;
        }
        $signature = (string)fread($handle, 8);
        fclose($handle);

        return $signature === "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
    }

    /**
     * TRUE if the file starts with the ZIP local file header signature
     *
     * An XLSX package is a ZIP archive, so it begins with "PK\x03\x04". This is
     * what lets open() tell a real XLSX apart from plain text, which is then
     * read as CSV.
     *
     * @param string $file
     *
     * @return bool
     */
    private static function isZip(string $file): bool
    {
        if (!is_readable($file) || is_dir($file)) {
            return false;
        }
        $handle = fopen($file, 'rb');
        if (!$handle) {
            return false;
        }
        $signature = (string)fread($handle, 4);
        fclose($handle);

        return in_array($signature, ["PK\x03\x04", "PK\x05\x06", "PK\x07\x08"], true);
    }

    /**
     * Open CSV through the common workbook API
     *
     * Returns CsvBook in 5.0. Use openCsvReader() for getCsvLine(), fromRow() and numeric CSV column keys.
     *
     * @param string $file
     * @param CsvOptions|array|null $options
     *
     * @return CsvBook
     */
    public static function openCsv(string $file, $options = []): CsvBook
    {
        return new CsvBook($file, $options);
    }

    /**
     * Check required XML workbook/ODS package parts and XML well-formedness
     *
     * This is not OOXML schema validation. The libxml error mode is restored; its diagnostic buffer is cleared before and after validation. XML errors are returned in $errors; missing package parts return false without XML errors.
     *
     * @param string $file
     * @param array|null $errors
     *
     * @return bool
     */
    public static function validate(string $file, ?array &$errors = []): bool
    {
        $errors = [];
        if (!is_file($file) || !is_readable($file)) {
            throw new Exception('Cannot read workbook: ' . $file);
        }
        $xmlReader = self::createReader($file);
        $previous = libxml_use_internal_errors(true);
        libxml_clear_errors();
        try {
            if (!self::isXlsx($file) && !self::isOds($file)) {
                $errors = libxml_get_errors();
                return false;
            }
            $fileList = $xmlReader->fileList();
            // Structural check only; this method does not validate OOXML schemas.
            foreach (self::isOds($file) ? ['mimetype', 'content.xml', 'META-INF/manifest.xml'] : ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels'] as $required) {
                if (!in_array($required, $fileList, true)) {
                    return false;
                }
            }
            $result = true;
            foreach ($fileList as $innerFile) {
                if (in_array(strtolower(pathinfo($innerFile, PATHINFO_EXTENSION)), ['xml', 'rels', 'vml'], true)) {
                    try {
                        if (!$xmlReader->openZip($innerFile)) {
                            $result = false;
                        }
                        else {
                            while ($xmlReader->read()) {
                                // Walk every XML part, including parts not used for reading cells.
                            }
                        }
                    }
                    finally {
                        $xmlReader->close();
                    }
                    if (libxml_get_errors()) {
                        $result = false;
                        $errors = array_merge($errors, libxml_get_errors());
                        libxml_clear_errors();
                    }
                }
            }

            return $result;
        }
        finally {
            try {
                $xmlReader->close();
            }
            finally {
                libxml_clear_errors();
                libxml_use_internal_errors($previous);
            }
        }
    }

    /**
     * Create sheet object
     *
     * @param string $sheetName
     * @param int|string $sheetId
     * @param string $file
     * @param string $path
     * @param XlsxBook $excel
     *
     * @return InterfaceSheetReader
     */
    public static function createSheet(string $sheetName, $sheetId, $file, $path, $excel): InterfaceSheetReader
    {
        return new Sheet($sheetName, $sheetId, $file, $path, $excel);
    }

    /**
     * Create XML reader object
     *
     * @param string $file
     * @param array|null $parserProperties
     *
     * @return InterfaceXmlReader
     */
    public static function createReader(string $file, ?array $parserProperties = []): InterfaceXmlReader
    {
        return new Reader($file, $parserProperties);
    }
}
