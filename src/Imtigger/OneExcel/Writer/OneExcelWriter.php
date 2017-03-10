<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\OneExcelWriterInterface;

abstract class OneExcelWriter implements OneExcelWriterInterface
{
    public static $format_supported = [];

    protected function isFormatSupported($format) {
        return in_array($format, static::$format_supported);
    }

    protected function checkFormatSupported($format) {
        if (!$this->isFormatSupported($format)) {
            throw new \Exception("Format {$format} is not supported by " . static::class);
        }
    }

    protected function getFormatMime($format) {
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