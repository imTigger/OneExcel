<?php
require __DIR__ . '/../vendor/autoload.php';

use Imtigger\OneExcel\OneExcelWriterFactory;
use Imtigger\OneExcel\Driver;
use Imtigger\OneExcel\Format;

$output = basename(__FILE__ . '.csv');
$excel = OneExcelWriterFactory::create()
    ->toStream($output, Format::CSV)
    ->withDriver(Driver::SPOUT)
    ->make();

for ($i = 1; $i < 1048576; $i += 1) {
    usleep(10); // Artificial delay
    $excel->writeRow($i, ['One', 'Excel', 'Test', $i]);
}

$excel->output();

