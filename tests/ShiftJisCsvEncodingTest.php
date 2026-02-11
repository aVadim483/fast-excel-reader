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
     * Этот тест проверяет “хороший” сценарий:
     * - CSV реально в Shift_JIS/CP932
     * - мы явно задаём inputEncoding
     * - openCsvStream() вешает iconv-фильтр, и дальше чтение идёт как UTF-8
     *
     * @requires extension iconv
     */
    public function testShiftJisCp932ExplicitInputEncodingIsConvertedToUtf8(): void
    {
        if (!function_exists('iconv')) {
            $this->markTestSkipped('iconv is required');
        }

        // UTF-8 источник (как должно быть прочитано на выходе)
        $utf8Lines = [
            "名前,年齢,都市\n",
            "山田太郎,30,東京\n",
        ];
        $utf8 = implode('', $utf8Lines);

        // В Windows-реальности чаще встречается CP932 (aka SJIS-win).
        // Для iconv корректнее использовать CP932 или SJIS-win.
        $cp932 = mb_convert_encoding($utf8, 'CP932', 'UTF-8');
        if ($cp932 === false || $cp932 === '') {
            $this->markTestSkipped('iconv cannot convert UTF-8 to CP932 on this system');
        }

        //$this->tmpFile = __DIR__ . '/test_files/sjis.csv';
        $file = __DIR__ . '/test_files/sjis.csv';
        file_put_contents($file, $cp932);

        // Важно: здесь мы явно говорим, что вход CP932.
        $csv = \avadim\FastExcelReader\Excel::openCsv($file, ['encoding' => 'CP932']);
        /*
        $opened = openCsvStream($this->tmpFile, [
            'inputEncoding' => 'CP932',
            'outputEncoding' => 'UTF-8',
            'detectEncoding' => false, // чтобы точно проверить работу фильтра, а не эвристики
            'stripBom' => true,
        ]);

        $fp = $opened['fp'];

        // Читаем уже как UTF-8 (после stream filter)
        $row1 = fgetcsv($fp, 0, ',', '"', ''); // escape отключаем, ближе к RFC
        $row2 = fgetcsv($fp, 0, ',', '"', '');

        fclose($fp);
        */
        $row1 = $csv->getCsvLine();
        $row2 = $csv->getCsvLine();

        $this->assertSame(['名前', '年齢', '都市'], $row1);
        $this->assertSame(['山田太郎', '30', '東京'], $row2);
    }

    /**
     * Этот тест показывает текущий лимит твоей эвристики:
     * без явного inputEncoding твоя guessInputEncodingFromSample()
     * НЕ умеет отличать японские кодировки (SJIS/CP932),
     * поэтому будет выбрана не та кодировка, и результат окажется “кракозябрами”.
     *
     * Если ты позже добавишь эвристику под Japanese encodings —
     * тогда этот тест надо будет обновить (ожидать корректное автоопределение).
     *
     * @requires extension iconv
     */
    public function testShiftJisAutoDetectCurrentlyDoesNotDetectCp932(): void
    {
        if (!function_exists('iconv')) {
            $this->markTestSkipped('iconv is required');
        }

        $utf8 = "名前,年齢,都市\n山田太郎,30,東京\n";
        $cp932 = @iconv('UTF-8', 'CP932//IGNORE', $utf8);
        if ($cp932 === false || $cp932 === '') {
            $this->markTestSkipped('iconv cannot convert UTF-8 to CP932 on this system');
        }

        $this->tmpFile = tempnam(sys_get_temp_dir(), 'csv_sjis_');
        file_put_contents($this->tmpFile, $cp932);

        // Здесь намеренно НЕ задаём inputEncoding — проверяем текущее поведение эвристики.
        $opened = openCsvStream($this->tmpFile, [
            'detectEncoding' => true,
            'outputEncoding' => 'UTF-8',
            'stripBom' => true,
        ]);

        $fp = $opened['fp'];
        $row1 = fgetcsv($fp, 0, ',', '"', '');
        fclose($fp);

        // Текущая реализация почти наверняка не угадает CP932,
        // значит заголовки НЕ совпадут с ожидаемыми японскими строками.
        $this->assertNotSame(['名前', '年齢', '都市'], $row1);
    }
}
