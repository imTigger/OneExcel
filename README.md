# OneExcel
PHP Excel read/write abstraction layer, support PHPExcel, LibXL and Spout

## Writer

### Basic Usages

```
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

```
$writer = OneExcelWriterFactory::create($output_format = OneExcelWriterInterface::FORMAT_XLSX)
```

```
$writer = OneExcelWriterFactory::createFromFile($filename, $output_format = OneExcelWriterInterface::FORMAT_XLSX, $input_format = OneExcelWriterInterface::FORMAT_AUTO)
``` 

#### OneExcelWriter

```
$writer->create($output_format = self::FORMAT_XLSX)
```

```
$writer->load($filename, $output_format = self::FORMAT_XLSX, $input_format = self::FORMAT_AUTO)
```

```
$writer->writeCell($row_num, $column_num, $data, $data_type = self::COLUMN_TYPE_STRING)
```

```
$writer->download($filename)
```

```
$writer->save($path)
```


## Reader

Not implemented yet

## Known Issues

- OneExcelWriterFactory do not auto create Spout driver
- Spout do not support random read/write rows (Won't fix)

## TODO

- [x] Register to [Packagist](https://packagist.org/packages/imtigger/oneexcel)
- [x] Emulate writeCell() behavior for Spout driver
- [ ] Implement $writer->writeRow($arr)
- [x] Implement COLUMN_TYPE_NUMERIC, COLUMN_TYPE_FORMULA for all drivers
- [ ] Implement COLUMN_TYPE_DATE, COLUMN_TYPE_TIME, COLUMN_TYPE_DATETIME for all drivers
- [ ] Implement COLUMN_TYPE_BOOLEAN, COLUMN_TYPE_NULL for LibXL driver
- [ ] Implement COLUMN_TYPE_* for Spout driver (Require upstream update)
- [ ] Implement OneExcelWriterFactory auto driver selection by format
- [ ] Implement sheet support
- [ ] Implement Reader
