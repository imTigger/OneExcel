<?php
namespace Imtigger\OneExcel;

use Imtigger\OneExcel\Writer\LibXLWriter;
use Imtigger\OneExcel\Writer\PHPExcelWriter;
use Imtigger\OneExcel\Writer\SpoutWriter;

class OneExcelWriterFactory
{
    public static function create($output_format = OneExcelWriterInterface::FORMAT_XLSX)
    {
        $driver = self::getDriver($output_format);
        $driver->create($output_format);
        return $driver;
    }

    public static function createFromFile($filename, $output_format = OneExcelWriterInterface::FORMAT_XLSX, $input_format = OneExcelWriterInterface::FORMAT_AUTO)
    {
        self::autoDetectInputFormat($filename, $input_format);
        $driver = self::getDriver($output_format, $input_format);
        $driver->load($filename, $output_format, $input_format);
        return $driver;
    }


    private static function autoDetectInputFormat($filename, &$input_format)
    {
        if ($input_format == OneExcelWriterInterface::FORMAT_AUTO) {
            $input_format = self::guessFormatFromFilename($filename);
        }
    }

    private static function guessFormatFromFilename($filename)
    {
        $pathinfo = pathinfo($filename);

        switch(strtolower($pathinfo['extension'])) {
            case 'csv':
                return OneExcelWriterInterface::FORMAT_CSV;
            case 'xls':
                return OneExcelWriterInterface::FORMAT_XLS;
            case 'xlsx':
                return OneExcelWriterInterface::FORMAT_XLSX;
            case 'ods':
                return OneExcelWriterInterface::FORMAT_ODS;
            default:
                throw new \Exception("Could not guess format for filename {$filename}");
        }
    }

    private static function getDriver($output_format, $input_format = null)
    {
        if (in_array($output_format, [OneExcelWriterInterface::FORMAT_XLSX, OneExcelWriterInterface::FORMAT_XLS])) {
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
        } else if (in_array($output_format, [OneExcelWriterInterface::FORMAT_CSV, OneExcelWriterInterface::FORMAT_ODS]) && $input_format == null) {
            return new SpoutWriter();
        } else {
            return new PHPExcelWriter();
        }
    }
}