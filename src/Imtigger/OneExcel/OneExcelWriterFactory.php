<?php
namespace Imtigger\OneExcel;

use Imtigger\OneExcel\Writer\FCsvWriter;
use Imtigger\OneExcel\Writer\LibXLWriter;
use Imtigger\OneExcel\Writer\OneExcelWriter;
use Imtigger\OneExcel\Writer\PhpSpreadsheetWriter;
use Imtigger\OneExcel\Writer\SpoutWriter;

class OneExcelWriterFactory
{
    private $driver = Driver::AUTO;
    private $input_format;
    private $output_format;
    private $output_mode;
    private $input_filename;
    private $output_filename;

    /**
     * @return OneExcelWriterFactory
     */
    public static function create()
    {
        return new OneExcelWriterFactory();
    }

    /**
     * @param $filename
     * @param string $input_format
     * @return $this
     */
    public function fromFile($filename, $input_format = Format::AUTO)
    {
        $this->input_filename = $filename;
        $this->input_format = $input_format;

        return $this;
    }

    /**
     * @param $driver
     * @return $this
     */
    public function withDriver($driver) {
        $this->driver = $driver;
        return $this;
    }

    /**
     * @param $filename
     * @param string $format
     * @return $this
     */
    public function toFile($filename, $format = Format::AUTO) {
        $this->output_filename = $filename;
        $this->output_mode = OneExcelWriter::OUTPUT_MODE_FILE;
        $this->output_format = $format;

        return $this;
    }

    /**
     * @param $filename
     * @param string $format
     * @return $this
     */
    public function toStream($filename, $format = Format::AUTO) {
        $this->output_filename = $filename;
        $this->output_mode = OneExcelWriter::OUTPUT_MODE_STREAM;
        $this->output_format = $format;

        return $this;
    }

    /**
     * @param $filename
     * @param string $format
     * @return $this
     */
    public function toDownload($filename, $format = Format::AUTO) {
        $this->output_filename = $filename;
        $this->output_mode = OneExcelWriter::OUTPUT_MODE_DOWNLOAD;
        $this->output_format = $format;

        return $this;
    }

    /**
     * @return OneExcelWriter
     * @throws \Exception
     */
    public function make() {
        if (!empty($this->input_filename)) {
            $this->autoDetectFormatFromFilename($this->input_format, $this->input_filename);
        }

        if (!empty($this->output_filename)) {
            $this->autoDetectFormatFromFilename($this->output_format, $this->output_filename);
        }

        if ($this->driver != Driver::AUTO) {
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

    /**
     * @param $input_format
     * @param $filename
     * @throws \Exception
     */
    private function autoDetectFormatFromFilename(&$input_format, $filename)
    {
        if ($input_format == Format::AUTO) {
            $input_format = $this->guessFormatFromFilename($filename);
        }
    }

    /**
     * @param $filename
     * @return string
     * @throws \Exception
     */
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

    /**
     * @param $driver
     * @return string
     * @throws \Exception
     */
    private function getDriverByName($driver) {
        switch ($driver) {
            case Driver::PHPEXCEL:
            case Driver::PHPSPREADSHEET:
                return PhpSpreadsheetWriter::class;
            case Driver::LIBXL:
                return LibXLWriter::class;
            case Driver::SPOUT:
                return SpoutWriter::class;
            case Driver::FPUTCSV:
                return FCsvWriter::class;
        }
        throw new \Exception("Unknown driver {$driver}");
    }

    /**
     * @param $output_format
     * @param null $input_format
     * @return string
     */
    private function getDriverByFormat($output_format, $input_format = null)
    {
        if (class_exists('ExcelBook') && in_array($output_format, [Format::XLSX, Format::XLS]) && ($input_format == $output_format || $input_format == null)) {
            // If LibXL exists, consider it first. LibXL support only when input format and output format are the same
            return LibXLWriter::class;
        }  else if (in_array($output_format, [Format::CSV, Format::ODS]) && in_array($input_format, [Format::XLSX, Format::CSV, Format::ODS, null])) {
            // If output is CSV/ODS, use Spout
            return SpoutWriter::class;
        } else {
            // Otherwise use PhpSpreadsheet
            return PhpSpreadsheetWriter::class;
        }
    }
}