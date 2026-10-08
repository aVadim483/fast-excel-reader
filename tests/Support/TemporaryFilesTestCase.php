<?php

namespace avadim\FastExcelReader\Tests\Support;

use avadim\FastExcelReader\Reader;
use PHPUnit\Framework\TestCase;

abstract class TemporaryFilesTestCase extends TestCase
{
    protected string $directory;

    protected function setUp(): void
    {
        parent::setUp();
        $this->directory = sys_get_temp_dir() . '/fxr45_' . uniqid('', true);
        mkdir($this->directory);
        Reader::setTempDir($this->directory);
    }

    protected function tearDown(): void
    {
        Reader::setTempDir();
        gc_collect_cycles();
        foreach (glob($this->directory . '/*') as $file) {
            unlink($file);
        }
        rmdir($this->directory);
        parent::tearDown();
    }

    protected function temporaryFile(string $content, string $name = 'input.csv'): string
    {
        $file = $this->directory . '/' . $name;
        file_put_contents($file, $content);
        return $file;
    }

    protected function workbook(array $parts = []): string
    {
        $file = $this->directory . '/book.xlsx';
        copy(__DIR__ . '/../test_files/standard-file.xlsx', $file);
        $zip = new \ZipArchive();
        $zip->open($file);
        foreach ($parts as $name => $content) {
            if ($content === null) {
                $zip->deleteName($name);
            }
            else {
                $zip->addFromString($name, $content);
            }
        }
        $zip->close();
        return $file;
    }
}
