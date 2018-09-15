<?php
require __DIR__ . '/../vendor/autoload.php';

use Imtigger\OneExcel\OneExcelWriterFactory;
use Imtigger\OneExcel\ColumnType;

$output = basename(__FILE__ . '.ods');
$excel = OneExcelWriterFactory::create()->toFile($output)->make();

$excel->writeCell(1, 0, 'Hello');
$excel->writeCell(2, 1, 'World');
$excel->writeCell(3, 2, 3.141592653, ColumnType::NUMERIC);
$excel->writeRow(4, ['One', 'Excel']);

$excel->output();

echo "File {$output} created using " . get_class($excel);

