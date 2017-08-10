<?php
namespace Imtigger\OneExcel\Reader;

use Imtigger\OneExcel\OneExcelReaderInterface;

abstract class OneExcelReader implements OneExcelReaderInterface
{
    public static $input_format_supported = [];

    protected $input_format;

    protected function isInputFormatSupported($format)
    {
        return in_array($format, static::$input_format_supported);
    }


    protected function checkFormatSupported($input_format)
    {
        if ($input_format != null && !$this->isInputFormatSupported($input_format)) {
            throw new \Exception("Input format {$input_format} is not supported by " . static::class);
        }
    }
}