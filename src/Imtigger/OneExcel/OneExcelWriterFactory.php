<?php
namespace Imtigger\OneExcel;

use Imtigger\OneExcel\Writer\FPutCsvWriter;
use Imtigger\OneExcel\Writer\LibXLWriter;
use Imtigger\OneExcel\Writer\OneExcelWriter;
use Imtigger\OneExcel\Writer\PHPExcelWriter;
use Imtigger\OneExcel\Writer\SpoutWriter;

class OneExcelWriterFactory
{
    private $driver;
    private $input_format;
    private $output_format;
    private $output_mode;
    private $input_filename;
    private $output_filename;

    public static function create()
    {
        return new OneExcelWriterFactory();
    }

    public function fromFile($filename, $input_format = Format::AUTO)
    {
        $this->input_filename = $filename;
        $this->input_format = $input_format;

        self::autoDetectFormatFromFilename($this->input_format, $filename);

        return $this;
    }

    public function withDriver($driver) {
        $this->driver = $driver;
        return $this;
    }

    public function toFile($filename, $format = Format::AUTO) {
        $this->output_filename = $filename;
        $this->output_mode = 'file';
        $this->output_format = $format;

        self::autoDetectFormatFromFilename($this->output_format, $filename);

        return $this;
    }

    public function toDownload($filename, $format = Format::AUTO) {
        $this->output_filename = $filename;
        $this->output_mode = 'download';
        $this->output_format = $format;

        self::autoDetectFormatFromFilename($this->output_format, $filename);

        return $this;
    }

    public function make() {
        if ($this->driver !== null) {
            $driver = $this->getDriverByName($this->driver);
        } else {
            $driver = $this->getDriverByFormat($this->output_format, $this->input_format);
        }

        /** @var OneExcelWriter $driver_impl */
        $driver_impl = new $driver;

        $driver_impl->setOutputMode($this->output_mode);
        $driver_impl->setOutputFilename($this->output_filename);

        if ($this->input_filename == null) {
            $driver_impl->create($this->output_format);
        } else {
            $driver_impl->load($this->input_filename, $this->output_format, $this->input_format);
        }

        return $driver_impl;
    }

    private function autoDetectFormatFromFilename(&$input_format, $filename)
    {
        if ($input_format == Format::AUTO) {
            $input_format = self::guessFormatFromFilename($filename);
        }
    }

    private function guessFormatFromFilename($filename)
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

    private function getDriverByName($driver) {
        switch ($driver) {
            case Driver::PHPEXCEL:
                return PHPExcelWriter::class;
            case Driver::LIBXL:
                return LibXLWriter::class;
            case Driver::SPOUT:
                return SpoutWriter::class;
            case Driver::FPUTCSV:
                return FPutCsvWriter::class;
        }
        throw new \Exception("Unknown driver {$driver}");
    }

    private function getDriverByFormat($output_format, $input_format = null)
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