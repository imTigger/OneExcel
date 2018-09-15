<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelWriterInterface;

abstract class OneExcelWriter implements OneExcelWriterInterface
{
    const OUTPUT_MODE_FILE = 'file';
    const OUTPUT_MODE_DOWNLOAD = 'download';
    const OUTPUT_MODE_STREAM = 'stream';

    public static $input_format_supported = [];
    public static $output_format_supported = [];
    public static $input_output_same_format;

    protected $input_format;
    protected $output_format;

    protected $output_mode = self::OUTPUT_MODE_FILE;
    protected $output_filename;

    protected function isInputFormatSupported($format)
    {
        return in_array($format, static::$input_format_supported);
    }


    protected function isOutputFormatSupported($format)
    {
        return in_array($format, static::$output_format_supported);
    }

    protected function checkFormatSupported($output_format, $input_format = null)
    {
        if (static::$input_output_same_format == true && $input_format != null && $input_format != $output_format) {
            throw new \Exception("Input format and output format needed to be the same {$input_format} for " . static::class);
        }

        if ($input_format != null && !$this->isInputFormatSupported($input_format)) {
            throw new \Exception("Input format {$input_format} is not supported by " . static::class);
        }

        if (!$this->isOutputFormatSupported($output_format)) {
            throw new \Exception("Output format {$output_format} is not supported by " . static::class);
        }
    }

    protected function getFormatMime($format)
    {
        switch ($format) {
            case Format::XLSX:
                return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
            case Format::XLS:
                return 'application/vnd.ms-excel';
            case Format::CSV:
                return 'text/csv';
            case Format::ODS:
                return 'application/vnd.oasis.opendocument.spreadsheet';
        }

        throw new \Exception("Unknown format {$format}");
    }
    
    public function setOutputMode($mode)
    {
        $this->output_mode = $mode;
    }

    public function setOutputFilename($output_filename)
    {
        $this->output_filename = $output_filename;
    }

    public function getOutputExtension()
    {
        return $this->output_format;
    }
}