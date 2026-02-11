<?php
declare(strict_types=1);

use PHPUnit\Framework\TestCase;

final class ShiftJisCsvEncodingTest extends TestCase
{
    private string $tmpFile;

    protected function tearDown(): void
    {
        if (isset($this->tmpFile) && is_file($this->tmpFile)) {
            @unlink($this->tmpFile);
        }
    }

    /**
     * Explicit encoding indication
     */
    public function testShiftJisCp932ExplicitInputEncodingIsConvertedToUtf8(): void
    {
        // UTF-8 source
        $utf8Lines = [
            "名前,年齢,都市\n",
            "山田太郎,30,東京\n",
        ];
        $utf8 = implode('', $utf8Lines);
        $cp932 = mb_convert_encoding($utf8, 'CP932', 'UTF-8');

        $file = __DIR__ . '/test_files/cp932.csv';
        file_put_contents($file, $cp932);

        $csv = \avadim\FastExcelReader\Excel::openCsv($file, ['encoding' => 'CP932']);
        $row1 = $csv->getCsvLine();
        $row2 = $csv->getCsvLine();

        $this->assertSame(['名前', '年齢', '都市'], $row1);
        $this->assertSame(['山田太郎', '30', '東京'], $row2);
    }

    /**
     * Automatic encoding detection
     */
    public function testShiftJisAutoDetectCurrentlyDoesNotDetectCp932(): void
    {
        $utf8 = "名前,年齢,都市\n山田太郎,30,東京\n";
        $cp932 = mb_convert_encoding($utf8, 'CP932', 'UTF-8');

        $file = __DIR__ . '/test_files/cp932.csv';
        file_put_contents($file, $cp932);

        $csv = \avadim\FastExcelReader\Excel::openCsv($file);
        $row1 = $csv->getCsvLine();

        $this->assertNotSame(['名前', '年齢', '都市'], $row1);
    }
}
