# OneExcel
PHP Excel read/write abstraction layer, support [PHPExcel](https://github.com/PHPOffice/PHPExcel), [LibXL](https://github.com/iliaal/php_excel) and [Spout](https://github.com/box/spout)

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
$excel->writeCell(3, 3, 3.141592653, OneExcelWriterInterface::COLUMN_TYPE_NUMERIC);

// $excel->save('example.xlsx'); // Save to disk
$excel->download('example.xlsx'); // Trigger download
```

### Documentations

#### OneExcelWriterFactory

```php
$writer = OneExcelWriterFactory::create($output_format = OneExcelWriterInterface::FORMAT_XLSX)
```

```php
$writer = OneExcelWriterFactory::createFromFile($filename, $output_format = OneExcelWriterInterface::FORMAT_XLSX, $input_format = OneExcelWriterInterface::FORMAT_AUTO)
``` 

#### OneExcelWriter

```php
$writer->create($output_format = self::FORMAT_XLSX)
```

```php
$writer->load($filename, $output_format = self::FORMAT_XLSX, $input_format = self::FORMAT_AUTO)
```

```php
$writer->writeCell($row_num, $column_num, $data, $data_type = self::COLUMN_TYPE_STRING)
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

## TODO

- [x] Register to [Packagist](https://packagist.org/packages/imtigger/oneexcel)
- [x] Emulate writeCell() behavior for Spout writer
- [x] OneExcelWriterFactory auto create writers base on input/output format
- [ ] Implement $writer->writeRow($arr)
- [x] Implement COLUMN_TYPE_NUMERIC, COLUMN_TYPE_FORMULA for all drivers
- [ ] Implement COLUMN_TYPE_DATE, COLUMN_TYPE_TIME, COLUMN_TYPE_DATETIME for all drivers
- [ ] Implement COLUMN_TYPE_BOOLEAN, COLUMN_TYPE_NULL for LibXL driver
- [ ] Implement COLUMN_TYPE_* for Spout driver (Require upstream update)
- [ ] Implement sheet support
- [ ] Implement Reader
