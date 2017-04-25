<?php
require __DIR__ . '/../vendor/autoload.php';

use Imtigger\OneExcel\OneExcelWriterFactory;
use Imtigger\OneExcel\ColumnType;
use Imtigger\OneExcel\Driver;
use Imtigger\OneExcel\Format;

$template = __DIR__ . '/../templates/template.xlsx';
$output = basename(__FILE__ . '.csv');
$excel = OneExcelWriterFactory::create()
    ->fromFile($template, Format::XLSX)
    ->toStream($output, Format::CSV)
    ->withDriver(Driver::SPOUT)
    ->make();

$excel->writeCell(2, 0, 'Hello');
$excel->writeCell(3, 1, 'World');
$excel->writeCell(4, 2, 3.141592653, ColumnType::NUMERIC);
$excel->writeRow(5, ['One', 'Excel']);

$excel->output();

