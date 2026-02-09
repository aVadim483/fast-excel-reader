<?php

final class CsvDataProvider
{
    /**
     * Data provider for CSV parsing tests (strict vs lenient).
     *
     * Each dataset:
     *  - input: string (may contain "\n" to emulate multi-line record)
     *  - expectedStrict: array|null  (null => expect exception)
     *  - expectedLenient: array|null (null => expect exception)
     *  - note: string
     */
    public static function provideCsvRecords(): array
    {
        return [
            'basic_1_unquoted' => [
                'input' => "a,b,c\n",
                'expectedStrict' => ['a','b','c'],
                'expectedLenient' => ['a','b','c'],
                'note' => 'Simple unquoted fields',
            ],
            'basic_2_empty_middle' => [
                'input' => "a,,c\n",
                'expectedStrict' => ['a','','c'],
                'expectedLenient' => ['a','','c'],
                'note' => 'Empty field between delimiters',
            ],
            'basic_3_empty_edges' => [
                'input' => ",b,\n",
                'expectedStrict' => ['','b',''],
                'expectedLenient' => ['','b',''],
                'note' => 'Empty first and last field',
            ],
            'quoted_1_simple' => [
                'input' => "\"a\",\"b\",\"c\"\n",
                'expectedStrict' => ['a','b','c'],
                'expectedLenient' => ['a','b','c'],
                'note' => 'Quoted fields',
            ],
            'quoted_2_delimiter_inside' => [
                'input' => "\"a,b\",c\n",
                'expectedStrict' => ['a,b','c'],
                'expectedLenient' => ['a,b','c'],
                'note' => 'Delimiter inside quotes',
            ],
            'quoted_3_escaped_quote' => [
                'input' => "\"a\"\"b\",c\n",
                'expectedStrict' => ['a"b','c'],
                'expectedLenient' => ['a"b','c'],
                'note' => 'Escaped quote by doubling ("")',
            ],

            'multiline_1_ok' => [
                'input' => "1,\"line1\nline2\",X\n",
                'expectedStrict' => ['1', "line1\nline2", 'X'],
                'expectedLenient' => ['1', "line1\nline2", 'X'],
                'note' => 'Newline inside quoted field',
            ],
            'multiline_2_eof_inside_quotes' => [
                // no closing quote + EOF
                'input' => "1,\"line1\nline2,X",
                'expectedStrict' => null,
                'expectedLenient' => ['1', "line1\nline2,X"],
                'note' => 'EOF inside quoted field: strict throws, lenient accepts',
            ],

            'junk_after_closing_quote_spaces' => [
                // This one depends on whether strict allows spaces after closing quote.
                // If your strict forbids spaces -> set expectedStrict=null
                // If your strict allows spaces -> set expectedStrict=['ok','X']
                'input' => "\"ok\" ,X\n",
                'expectedStrict' => null,
                'expectedLenient' => ['ok','X'],
                'note' => 'Spaces after closing quote (set expectation per your policy)',
            ],
            'junk_after_closing_quote_append' => [
                'input' => "\"ok\"zzz,X\n",
                'expectedStrict' => null,
                'expectedLenient' => ['okzzz','X'],
                'note' => 'Junk after closing quote: strict throws, lenient appends to value',
            ],

            'quote_in_unquoted_1' => [
                'input' => "a\"b\"c,X\n",
                'expectedStrict' => null,
                'expectedLenient' => ['a"b"c','X'],
                'note' => 'Quote inside unquoted field: strict throws, lenient treats as char',
            ],
            'quote_in_unquoted_2_double_quotes' => [
                'input' => "a\"\"b,X\n",
                'expectedStrict' => null,
                'expectedLenient' => ['a""b','X'],
                'note' => 'Double quotes in unquoted: strict throws, lenient treats as chars',
            ],

            'unclosed_quote_eol' => [
                'input' => "\"abc,X\n",
                'expectedStrict' => null,
                'expectedLenient' => ['abc,X'],
                'note' => 'Unclosed quote before EOL: strict throws, lenient accepts',
            ],

            'triple_quotes_value_quote' => [
                'input' => "\"\"\"\",X\n",
                'expectedStrict' => ['"', 'X'],
                'expectedLenient' => ['"', 'X'],
                'note' => 'Field value is one quote (""")',
            ],
            'four_quotes_value_two_quotes' => [
                'input' => "\"\"\"\"\"\",X\n", // 6 quotes total: " "" "" " -> value: ""
                'expectedStrict' => ['""', 'X'],
                'expectedLenient' => ['""', 'X'],
                'note' => 'Field value is two quotes ("""""" => "")',
            ],

            'variable_columns_less' => [
                'input' => "a,b\n",
                // Strict: if you enforce expected width (e.g., from header), set null.
                // If strict does not enforce width, set ['a','b'].
                'expectedStrict' => ['a','b'],
                'expectedLenient' => ['a','b'],
                'note' => 'Fewer columns (width enforcement is optional, adjust expectedStrict)',
            ],
            'variable_columns_more' => [
                'input' => "a,b,c,d\n",
                'expectedStrict' => ['a','b','c','d'],
                'expectedLenient' => ['a','b','c','d'],
                'note' => 'More columns (width enforcement is optional, adjust expectedStrict)',
            ],

            'empty_line' => [
                'input' => "\n",
                // Strict can be [''] or null depending on your policy; lenient often skips empty lines.
                'expectedStrict' => [''],
                'expectedLenient' => [''],
                'note' => 'Empty line (adjust depending on skipEmptyLines option)',
            ],
            'crlf_line_ending' => [
                'input' => "a,b,c\r\n",
                'expectedStrict' => ['a','b','c'],
                'expectedLenient' => ['a','b','c'],
                'note' => 'CRLF line ending',
            ],
            'utf8_bom' => [
                'input' => "\xEF\xBB\xBFa,b\n",
                // If stripBom=true, expected is ['a','b']. Otherwise first field includes BOM.
                'expectedStrict' => ['a','b'],
                'expectedLenient' => ['a','b'],
                'note' => 'UTF-8 BOM at start (assumes stripBom=true)',
            ],
        ];
    }
}
