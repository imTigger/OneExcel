<?php
namespace Imtigger\OneExcel;

use Imtigger\OneExcel\Writer\LibXLWriter;
use Imtigger\OneExcel\Writer\PHPExcelWriter;

class OneExcelWriterFactory
{
    public static function create($output_format = OneExcelWriterInterface::FORMAT_XLSX)
    {
        $driver = self::getDriver();
        $driver->create($output_format);
        return $driver;
    }

    public static function createFromFile($filename, $output_format = OneExcelWriterInterface::FORMAT_XLSX, $input_format = OneExcelWriterInterface::FORMAT_AUTO)
    {
        $driver = self::getDriver();
        $driver->load($filename, $output_format, $input_format);
        return $driver;
    }

    public static function getDriver()
    {
        if (class_exists('ExcelBook')) {
            return new LibXLWriter();
        } else {
            return new PHPExcelWriter();
        }
    }
}