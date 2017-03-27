<?php
namespace Imtigger\OneExcel;

use Imtigger\OneExcel\Writer\FPutCsvWriter;
use Imtigger\OneExcel\Writer\LibXLWriter;
use Imtigger\OneExcel\Writer\PHPExcelWriter;
use Imtigger\OneExcel\Writer\SpoutWriter;

class OneExcelWriterFactory
{
    public static function create($format = Format::XLSX, $driverName = Driver::AUTO)
    {
        if ($driverName != Driver::AUTO) {
            $driver = self::getDriverByName($driverName);
        } else {
            $driver = self::getDriverByFormat($format);
        }

        $driver->create($format);

        return $driver;
    }

    public static function createFromFile($filename, $output_format = Format::XLSX, $input_format = Format::AUTO, $driverName = Driver::AUTO)
    {
        if ($driverName != Driver::AUTO) {
            $driver = self::getDriverByName($driverName);
        } else {
            self::autoDetectInputFormat($filename, $input_format);
            $driver = self::getDriverByFormat($output_format, $input_format);
        }
        $driver->load($filename, $output_format, $input_format);
        return $driver;
    }

    private static function autoDetectInputFormat($filename, &$input_format)
    {
        if ($input_format == Format::AUTO) {
            $input_format = self::guessFormatFromFilename($filename);
        }
    }

    private static function guessFormatFromFilename($filename)
    {
        $pathinfo = pathinfo($filename);

        switch(strtolower($pathinfo['extension'])) {
            case 'csv':
                return Format::CSV;
            case 'xls':
                return Format::XLS;
            case 'xlsx':
                return Format::XLSX;
            case 'ods':
                return Format::ODS;
            default:
                throw new \Exception("Could not guess format for filename {$filename}");
        }
    }

    private static function getDriverByName($driver) {
        switch ($driver) {
            case Driver::PHPEXCEL:
                return new PHPExcelWriter();
            case Driver::LIBXL:
                return new LibXLWriter();
            case Driver::SPOUT:
                return new SpoutWriter();
            case Driver::FPUTCSV:
                return new FPutCsvWriter();
        }
        throw new \Exception("Unknown driver {$driver}");
    }

    private static function getDriverByFormat($output_format, $input_format = null)
    {
        if (in_array($output_format, [Format::XLSX, Format::XLS])) {
            // If LibXL exists, consider it first
            if (class_exists('ExcelBook')) {
                // LibXL support only when input format and output format are the same
                if ($input_format == null || $input_format == $output_format) {
                    return new LibXLWriter();
                } else {
                    return new PHPExcelWriter();
                }
            } else {
                return new PHPExcelWriter();
            }
        } else if (in_array($output_format, [Format::CSV, Format::ODS]) && $input_format == null) {
            return new SpoutWriter();
        } else {
            return new PHPExcelWriter();
        }
    }
}