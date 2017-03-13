<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\OneExcelWriterInterface;

abstract class OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [];
    public static $output_format_supported = [];
    public static $input_output_same_format;

    protected function isInputFormatSupported($format)
    {
        return in_array($format, static::$input_format_supported);
    }


    protected function isOutputFormatSupported($format)
    {
        return in_array($format, static::$output_format_supported);
    }

    protected function checkFormatSupported($output_format, $input_format)
    {
        if (static::$input_output_same_format == true && $input_format != null && $input_format != $output_format) {
            throw new \Exception("Input format and output format needed to be the same {$input_format} for " . static::class);
        }

        if (!$this->isInputFormatSupported($input_format)) {
            throw new \Exception("Input format {$input_format} is not supported by " . static::class);
        }

        if (!$this->isOutputFormatSupported($output_format)) {
            throw new \Exception("Output format {$output_format} is not supported by " . static::class);
        }
    }

    protected function autoDetectInputFormat($filename, &$input_format)
    {
        if ($input_format == self::FORMAT_AUTO) {
            $input_format = self::guessFormatFromFilename($filename);
        }
    }

    protected function guessFormatFromFilename($filename)
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
                throw new Exception("Could not guess format for filename {$filename}");
        }
    }

    protected function getFormatMime($format)
    {
        switch ($format) {
            case self::FORMAT_XLSX:
                return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
            case self::FORMAT_XLS:
                return 'application/vnd.ms-excel';
            case self::FORMAT_CSV:
                return 'text/csv';
            case self::FORMAT_ODS:
                return 'application/vnd.oasis.opendocument.spreadsheet';
        }
    }
}