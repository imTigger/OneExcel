# OneExcel
PHP Excel read/write abstraction layer, support [PHPExcel](https://github.com/PHPOffice/PHPExcel), [LibXL](https://github.com/iliaal/php_excel), [Spout](https://github.com/box/spout) and PHP `fputcsv`

Targets to simplify server compatibility issue between Excel libraries.

Ideal for simple-formatted but huge spreadsheet files

## Installation

### Requirements

- PHP 5.5+
- `php_zip`, `php_xmlreader`, `php_simplexml` enabled
- (Recommended) LibXL installed & `php_excel` enabled

### Composer

OneExcel can only be installed from [Composer](https://getcomposer.org/).

Run the following command:
```
$ composer require imtigger/oneexcel
```

## Writer

### Basic Usages

```php
// $excel = OneExcelWriterFactory::create(); // Create Excel from scratch
$excel = OneExcelWriterFactory::createFromFile('templates/manifest.xlsx'); // Create Excel from template

$excel->writeCell(1, 1, 'Hello');
$excel->writeCell(2, 2, 'World');
$excel->writeCell(3, 3, 3.141592653, ColumnType::NUMERIC);
$excel->writeRow(4, ['One', 'Excel']);

// $excel->save('example.xlsx'); // Save to disk
$excel->download('example.xlsx'); // Trigger download
```

### Documentations

#### OneExcelWriterFactory

```php
$writer = OneExcelWriterFactory::create($output_format = Format::XLSX)
```

```php
$writer = OneExcelWriterFactory::createFromFile($filename, $output_format = Format::XLSX, $input_format = Format::AUTO)
```

#### OneExcelWriter

```php
$writer->create($output_format = Format::XLSX)
```

```php
$writer->load($filename, $output_format = Format::XLSX, $input_format = Format::AUTO)
```

```php
$writer->writeCell($row_num, $column_num, $data, $data_type = ColumnType::STRING)
```

```php
$writer->writeRow($row_num, $data)
```

```php
$writer->download($filename)
```

```php
$writer->save($path)
```


## Reader

Not implemented yet

## Known Issues

- Spout do not support random read/write rows (Upstream limitation, Won't fix)
- Spout do not support formula (Upstream limitation, Won't fix)
- fputcsv driver ignores all ColumnType::* (File-type limitation, Won't fix)

## TODO

- [x] Register to [Packagist](https://packagist.org/packages/imtigger/oneexcel)
- [x] Emulate writeCell() behavior for Spout/fuptcsv writer
- [x] OneExcelWriterFactory auto create writers base on input/output format
- [x] Refactor: Move constants to separate class
- [x] Implement $writer->writeRow($arr)
- [x] Implement ColumnType::NUMERIC, ColumnType::FORMULA for all drivers
- [ ] Implement ColumnType::DATE, ColumnType::TIME, ColumnType::DATETIME for all drivers
- [ ] Implement ColumnType::BOOLEAN, ColumnType::NULL for LibXL driver
- [ ] Implement ColumnType::* for Spout driver (Require upstream update)
- [ ] Implement sheet support
- [ ] Implement Reader
