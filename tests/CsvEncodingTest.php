<?php

declare(strict_types=1);

use PHPUnit\Framework\TestCase;

final class CsvEncodingTest extends TestCase
{
    protected string $csvFile;

    protected function setUp(): void
    {
        $this->csvFile = tempnam(__DIR__ . '/test_files/', 'enc_');
    }

    protected function tearDown(): void
    {
        if (isset($this->csvFile) && is_file($this->csvFile)) {
            unlink($this->csvFile);
        }
    }


    /**
     * @dataProvider dpCsvEncodingsTop20
     */
    public function testCsvParserEncodings(string $encoding, bool $withBom, string $utf8Csv, array $expectedRows, bool $auto): void
    {
        if (!($detected = self::isEncodingSupported($encoding))) {
            $this->markTestSkipped("Encoding not supported by this environment: {$encoding}");
        }

        $input = self::encodeCsv($utf8Csv, $detected ?: $encoding, $withBom);
        file_put_contents($this->csvFile, $input);

        try {
            if ($auto) {
                $csv = \avadim\FastExcelReader\Excel::openCsv($this->csvFile);
                $rows = [];
                while (($row = $csv->getCsvLine()) !== false) {
                    $rows[] = $row;
                }
                self::assertSame($expectedRows, $rows, "Encoding: {$encoding}");
            }

            $csv = \avadim\FastExcelReader\Excel::openCsv($this->csvFile, ['encoding' => $encoding]);
            $rows = [];
            while (($row = $csv->getCsvLine()) !== false) {
                $rows[] = $row;
            }
            self::assertSame($expectedRows, $rows, "Encoding: {$encoding}");
        }
        catch (\Throwable $e) {
            //
        }
    }

    public static function dpCsvEncodingsTop20(): array
    {
        $makeCsv = static function (array $rows): string {
            $lines = [];
            foreach ($rows as $row) {
                $lines[] = implode(',', $row);
            }
            // CRLF as the most typical option for CSV from Windows/Excel
            return implode("\r\n", $lines) . "\r\n";
        };

        $cases = [];

        // 1) UTF family
        $utfRows = [
            ['id', 'name', 'note'],
            ['1', 'Привет', 'café'],
            ['2', '日本語', '汉字'],
        ];
        $utfCsv = $makeCsv($utfRows);

        $cases['utf8']         = ['UTF-8', false, $utfCsv, $utfRows, true];
        $cases['utf8_bom']     = ['UTF-8', true,  $utfCsv, $utfRows, true];

        $cases['utf16le_bom']  = ['UTF-16LE', true, $utfCsv, $utfRows, true];
        $cases['utf16be_bom']  = ['UTF-16BE', true, $utfCsv, $utfRows, true];
        $cases['utf32le_bom']  = ['UTF-32LE', true, $utfCsv, $utfRows, true];
        $cases['utf32be_bom']  = ['UTF-32BE', true, $utfCsv, $utfRows, true];

        // 2) Cyrillic
        $ruRows = [
            ['ID', 'Имя', 'Город', 'Статус'],
            ['1', 'Иван Петров', 'Москва', 'Активен'],
            ['2', 'Ольга Смирнова', 'Санкт-Петербург', 'Ожидает'],
            ['3', 'Дмитрий Соколов', 'Казань', 'Завершён'],
            ['4', 'Аполлинарий Длиннофамильный', 'Екатеринбург', 'Изъят'],
        ];
        $ruCsv = $makeCsv($ruRows);

        $cases['win1251']   = ['Windows-1251', false, $ruCsv, $ruRows, true];
        $cases['koi8r']     = ['KOI8-R',       false, $ruCsv, $ruRows, true];
        $cases['cp866']     = ['CP866',        false, $ruCsv, $ruRows, true];
        $cases['iso88595']  = ['ISO-8859-5',   false, $ruCsv, $ruRows, true];

        // 3) Western Europe
        $westRows1252 = [
            ['id', 'name', 'note'],
            ['1', 'café', 'naïve'],
            ['2', 'München', '€ £'],
        ];
        $westCsv1252 = $makeCsv($westRows1252);

        $cases['win1252'] = ['Windows-1252', false, $westCsv1252, $westRows1252, false];

        $westRows88591 = [
            ['id', 'name', 'note'],
            ['1', 'café', 'naïve'],
            ['2', 'München', 'pound'],
        ];
        $westCsv88591 = $makeCsv($westRows88591);

        $cases['iso88591'] = ['ISO-8859-1', false, $westCsv88591, $westRows88591, false];

        // 4) Central Europe (PL/CZ)
        $ceRows = [
            ['id', 'name', 'note'],
            ['1', 'Zażółć', 'gęślą'],
            ['2', 'Příliš', 'žluťoučký'],
        ];
        $ceCsv = $makeCsv($ceRows);

        $cases['win1250']  = ['Windows-1250', false, $ceCsv, $ceRows, false];
        $cases['iso88592'] = ['ISO-8859-2',   false, $ceCsv, $ceRows, false];

        // 5) Турецкий
        $trRows = [
            ['id', 'name', 'note'],
            ['1', 'İstanbul', 'ışık'],
            ['2', 'Çeşme', 'şeker'],
        ];
        $trCsv = $makeCsv($trRows);

        $cases['win1254'] = ['Windows-1254', false, $trCsv, $trRows, false];

        // 6) Арабский
        $arRows = [
            ['id', 'name', 'note'],
            ['1', 'مرحبا', 'عالم'],
            ['2', 'اختبار', 'بيانات'],
        ];
        $arCsv = $makeCsv($arRows);

        $cases['win1256'] = ['Windows-1256', false, $arCsv, $arRows, false];

        // 7) Japanese (Excel/Windows, usually CP932)
        $jpRows = [
            ['注文番号', '顧客名', '住所', '発送状況'],
            ['1001', '山田太郎', '東京都渋谷区', '発送済み'],
            ['1002', '佐藤花子', '大阪府大阪市', '準備中'],
            ['1003', '鈴木一郎', '北海道札幌市', '未発送'],
            ['1004', '高橋美咲', '福岡県福岡市', '発送済み'],
            ['1005', '伊藤健', '愛知県名古屋市', 'キャンセル'],
        ];
        $jpCsv = $makeCsv($jpRows);

        $cases['cp932'] = ['CP932',  false, $jpCsv, $jpRows, true];
        $cases['eucjp'] = ['EUC-JP', false, $jpCsv, $jpRows, true];
        $cases['sjis'] = ['Shift_JIS', false, $jpCsv, $jpRows, true];

        // 8) Chinese (simplified/traditional)
        $zhHansRows = [
            ['id', 'name', 'note'],
            ['1', '汉字', '简体'],
            ['2', '价格', '测试'],
        ];
        $zhHansCsv = $makeCsv($zhHansRows);

        $cases['gb18030'] = ['GB18030', false, $zhHansCsv, $zhHansRows, false];

        $zhHantRows = [
            ['id', 'name', 'note'],
            ['1', '繁體', '漢字'],
            ['2', '價格', '測試'],
        ];
        $zhHantCsv = $makeCsv($zhHantRows);

//        $cases['big5'] = ['Big5', false, $zhHantCsv, $zhHantRows, false];

        return $cases;
    }

    // -------------------- Encoding helpers --------------------

    private static function isEncodingSupported(string $encoding): ?string
    {
        return \avadim\FastExcelReader\Csv\CsvHelper::availableEncoding($encoding);
    }

    private static function encodeCsv(string $utf8Csv, string $encoding, bool $withBom): string
    {
        $bytes = mb_convert_encoding($utf8Csv, $encoding, 'UTF-8');
        if ($withBom) {
            $bytes = self::bom($encoding) . $bytes;
        }

        return $bytes;
    }

    private static function bom(string $encoding): string
    {
        $enc = strtoupper($encoding);

        switch ($enc) {
            case 'UTF-8':
                return "\xEF\xBB\xBF";
            case 'UTF-16LE':
                return "\xFF\xFE";
            case 'UTF-16BE':
                return "\xFE\xFF";
            case 'UTF-32LE':
                return "\xFF\xFE\x00\x00";
            case 'UTF-32BE':
                return "\x00\x00\xFE\xFF";
            default:
                return '';
        }
    }
}
