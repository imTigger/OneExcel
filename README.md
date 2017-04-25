# OneExcel

[![Latest Stable Version](https://poser.pugx.org/imtigger/oneexcel/v/stable)](https://packagist.org/packages/imtigger/oneexcel)
[![Latest Unstable Version](https://poser.pugx.org/imtigger/oneexcel/v/unstable)](https://packagist.org/packages/imtigger/oneexcel)
[![Total Downloads](https://poser.pugx.org/imtigger/oneexcel/downloads)](https://packagist.org/packages/imtigger/oneexcel)
[![License](https://poser.pugx.org/imtigger/oneexcel/license)](https://packagist.org/packages/imtigger/oneexcel)

PHP Excel read/write abstraction layer, support [PHPExcel](https://github.com/PHPOffice/PHPExcel), [LibXL](https://github.com/iliaal/php_excel), [Spout](https://github.com/box/spout) and PHP `fputcsv`

Targets to simplify server compatibility issue between Excel libraries.

Ideal for simple-formatted but huge spreadsheet files

## Installation

### Requirements

- PHP >= 5.6
- `php_zip`, `php_xmlreader`, `php_simplexml` enabled
- (Recommended) LibXL installed & `php_excel` enabled

### Composer

OneExcel can only be installed from [Composer](https://getcomposer.org/).

Run the following command:
```
$ composer require imtigger/oneexcel
```

## Writer

### Documentations

#### Basic Usage

```php
$excel = OneExcelWriterFactory::create()
        ->create()
        ->toFile('excel.xlsx')
        ->make();
        
$excel->writeCell(1, 0, 'Hello');
$excel->writeCell(2, 1, 'World');
$excel->writeCell(3, 2, 3.141592653, ColumnType::NUMERIC);
$excel->writeRow(4, ['One', 'Excel']);

$excel->output();
```

#### Advanced Usage

```php
$excel = OneExcelWriterFactory::create()
        ->create()
        ->fromFile('template.xlsx', Format::XLSX)
        ->toFile('excel.xlsx', Format::CSV)
        ->withDriver(Driver::SPOUT)
        ->make();
        
$excel->writeCell(1, 0, 'Hello');
$excel->writeCell(2, 1, 'World');
$excel->writeCell(3, 2, 3.141592653, ColumnType::NUMERIC);
$excel->writeRow(4, ['One', 'Excel']);

$excel->output();
```

## Reader

Not implemented yet

## Known Issues

- Spout do not support random read/write rows (Upstream limitation, Won't fix)
- Spout do not support formula (Upstream limitation, Won't fix)
- fputcsv driver ignores all ColumnType::* (File-type limitation, Won't fix)

## TODO

- [x] Register to [Packagist](https://packagist.org/packages/imtigger/oneexcel)
- [x] Emulate writeCell() behavior for Spout/fputcsv writer
- [x] OneExcelWriterFactory auto create writers base on input/output format
- [x] Refactor: Move constants to separate class
- [x] Implement load() for SpoutWriter and FPutCsvWriter
- [x] Implement $writer->writeRow($arr)
- [x] Implement ColumnType::NUMERIC, ColumnType::FORMULA for all drivers
- [ ] Implement ColumnType::DATE, ColumnType::TIME, ColumnType::DATETIME for all drivers
- [ ] Implement ColumnType::BOOLEAN, ColumnType::NULL for LibXL driver
- [ ] Implement ColumnType::* for Spout driver (Require upstream update)
- [ ] Implement sheet support
- [ ] Implement Reader
