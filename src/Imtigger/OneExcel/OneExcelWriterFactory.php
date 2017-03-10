<?php
namespace Imtigger\OneExcel;

use Imtigger\OneExcel\Writer\LibXLWriter;
use Imtigger\OneExcel\Writer\PHPExcelWriter;

class OneExcelWriterFactory
{
    public static function create($type = OneExcelWriterInterface::FORMAT_XLSX) {
        if (class_exists('ExcelBook')) {
            return new LibXLWriter();
        } else {
            return new PHPExcelWriter();
        }
    }
}