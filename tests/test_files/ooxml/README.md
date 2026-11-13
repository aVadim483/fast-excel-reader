# OOXML family fixtures / Фикстуры XML-семейства

Created on 2026-11-13 with Microsoft Excel on Windows, with AutomationSecurity=3
(macros disabled). All four files were saved by Excel in their declared formats;
they are not renamed XLSX files. ReaderData contains ID/Name headers, numeric 1,
Unicode text, numeric 45000, cached formula result 3, boolean true and merged A4:B4.
Second contains “Second sheet”. XLSM and XLTM contain xl/vbaProject.bin.

Созданы 2026-11-13 в Microsoft Excel на Windows при AutomationSecurity=3
(макросы отключены). Все четыре файла сохранены Excel в соответствующем формате,
а не переименованы. ReaderData содержит заголовки ID/Name, число 1, Unicode-текст,
число 45000, сохранённый результат формулы 3, boolean true и объединение A4:B4.
Second содержит «Second sheet». XLSM и XLTM содержат xl/vbaProject.bin.

The VBA binary was taken from the PhpSpreadsheet MIT-licensed fixture below,
attached to the sample workbook, then resaved by Excel. It is retained solely to
verify reading a workbook containing a real VBA project. VBA is not executed by
fixture preparation or tests. License: PHPSPREADSHEET-LICENSE.txt.

VBA-проект взят из MIT-фикстуры PhpSpreadsheet по ссылке ниже, добавлен в тестовую
книгу и пересохранён Excel. Он нужен только для проверки чтения книги с настоящим
VBA-проектом. Код не выполняется при подготовке фикстур и тестах.
Лицензия: PHPSPREADSHEET-LICENSE.txt.

Source / Источник:
https://github.com/PHPOffice/PhpSpreadsheet/blob/master/tests/data/Writer/XLSX/form_pass_print.xlsm
