<?php
require __DIR__ . '/../vendor/autoload.php';

use Imtigger\OneExcel\OneExcelWriterFactory;
use Imtigger\OneExcel\ColumnType;
use Imtigger\OneExcel\Driver;
use Imtigger\OneExcel\Format;

// $excel = OneExcelWriterFactory::create(); // Create empty sheet
// $excel = OneExcelWriterFactory::create(Format::CSV); // Create empty sheet using specifed format
// $excel = OneExcelWriterFactory::create(Format::CSV, Driver::FPUTCSV); // Create empty sheet using specifed driver
//$excel = OneExcelWriterFactory::createFromFile('templates/template.xlsx'); // Create sheet using template
$excel = OneExcelWriterFactory::createFromFile('templates/template.csv', Format::CSV, Format::CSV, Driver::PHPEXCEL); // Create sheet using template
// $excel = OneExcelWriterFactory::createFromFile('templates/template.xlsx', Format::XLSX, Format::XLSX, Driver::LIBXL); // Create Excel from template specifing input/output format

$excel->writeCell(1, 0, 'Hello');
$excel->writeCell(2, 1, 'World');
$excel->writeCell(3, 2, 3.141592653, ColumnType::NUMERIC);
$excel->writeRow(4, ['One', 'Excel']);

// $excel->save('example.xlsx'); // Save to disk
$excel->save('example.csv'); // Trigger download

