<?php
namespace Imtigger\OneExcel;


class OneExcelWriterFactory
{
    public static function create($type = OneExcelWriter::FORMAT_XLSX) {
        if (class_exists('ExcelBook')) {
            return new LibXLWriter();
        } else {
            return new PHPExcelWriter();
        }
    }
}