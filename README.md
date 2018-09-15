# OneExcel

[![Build Status](https://travis-ci.org/imTigger/OneExcel.svg?branch=master)](https://travis-ci.org/imTigger/OneExcel)
[![Latest Stable Version](https://poser.pugx.org/imtigger/oneexcel/v/stable)](https://packagist.org/packages/imtigger/oneexcel)
[![Latest Unstable Version](https://poser.pugx.org/imtigger/oneexcel/v/unstable)](https://packagist.org/packages/imtigger/oneexcel)
[![Total Downloads](https://poser.pugx.org/imtigger/oneexcel/downloads)](https://packagist.org/packages/imtigger/oneexcel)
[![License](https://poser.pugx.org/imtigger/oneexcel/license)](https://packagist.org/packages/imtigger/oneexcel)

PHP Excel read/write abstraction layer, support [PhpSpreadSheet](https://github.com/PHPOffice/PhpSpreadSheet), [LibXL](https://github.com/iliaal/php_excel), [Spout](https://github.com/box/spout) and PHP `fputcsv`/`fgetcsv`

Targets to simplify server compatibility issue between Excel libraries and performance issue in huge files.

Ideal for simple-formatted but huge spreadsheet files such as reporting.

## Installation

### Requirements

- PHP >= 5.6.4
- `php_zip`, `php_xmlreader`, `php_simplexml` enabled
- (Recommended) LibXL installed & `php_excel` enabled

### Composer

OneExcel can only be installed from [Composer](https://getcomposer.org/).

Run the following command:
```
$ composer require imtigger/oneexcel
```

## Writer

### Basic Usage

```php
$excel = OneExcelWriterFactory::create()
        ->toFile('excel.xlsx')
        ->make();
        
$excel->writeCell(1, 0, 'Hello');
$excel->writeCell(2, 1, 'World');
$excel->writeCell(3, 2, 3.141592653, ColumnType::NUMERIC);
$excel->writeRow(4, ['One', 'Excel']);
$excel->writeCell(4, 2, 'Test');

$excel->output();
```

### Advanced Usage

```php
$excel = OneExcelWriterFactory::create()
        ->fromFile('template.xlsx', Format::XLSX)
        ->toStream('excel.csv', Format::CSV)
        ->withDriver(Driver::SPOUT)
        ->make();
        
$excel->writeCell(1, 0, 'Hello');
$excel->writeCell(2, 1, 'World');
$excel->writeCell(3, 2, 3.141592653, ColumnType::NUMERIC);
$excel->writeRow(4, ['One', 'Excel']);
$excel->writeCell(4, 2, 'Test');

$excel->output();
```

### Writer Default Driver

#### With Template

|      | XLSX  | XLS   | ODS   | CSV   | From |
|------|-------|-------|-------|-------|------|
| XLSX | LibXL | PhpSS | Spout | PhpSS |      | 
| XLS  | PhpSS | LibXL | -     | PhpSS |      |
| ODS  | PhpSS | PhpSS | Spout | PhpSS |      |
| CSV  | Spout | PhpSS | Spout | Spout |      |
| To   |       |       |       |       |      |

##### Without Template

|      | Driver | 
|------|--------|
| XLSX | Spout  | 
| XLS  | LibXL  |
| ODS  | Spout  |
| CSV  | Spout  |

* LibXL fallbacks to PhpSpreadSheet if not present

## Reader

(Version 0.6+)

```php
$excel = OneExcelReaderFactory::create()
        ->fromFile('excel.xlsx')
        // ->withDriver(Driver::SPOUT)
        ->make();
        
foreach ($excel->row() as $row) {
    //
}

$excel->close();
```

## Known Issues

- Spout reader driver output empty rows as SINGLE column (Upstream problem?)
- Spout do not support random read/write rows (Upstream limitation, Won't fix)
- Spout do not support formula (Upstream limitation, Won't fix)
- fputcsv driver ignores all ColumnType::* (File-type limitation, Won't fix)
