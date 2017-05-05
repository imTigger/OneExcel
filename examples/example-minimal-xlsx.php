<?php
require __DIR__ . '/../vendor/autoload.php';

use Imtigger\OneExcel\OneExcelWriterFactory;
use Imtigger\OneExcel\ColumnType;
use Imtigger\OneExcel\Driver;
use Imtigger\OneExcel\Format;

$output = basename(__FILE__ . '.xlsx');
$excel = OneExcelWriterFactory::create()->toFile($output)->make();

$excel->writeCell(1, 0, 'Hello');
$excel->writeCell(2, 1, 'World');
$excel->writeCell(3, 2, 3.141592653, ColumnType::NUMERIC);
$excel->writeRow(4, ['One', 'Excel']);

$excel->output();

echo "File {$output} created using " . get_class($excel);

