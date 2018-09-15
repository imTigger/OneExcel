<?php
namespace Imtigger\OneExcel;

use Imtigger\OneExcel\Reader\FCsvReader;
use Imtigger\OneExcel\Reader\LibXLReader;
use Imtigger\OneExcel\Reader\OneExcelReader;
use Imtigger\OneExcel\Reader\PHPExcelReader;
use Imtigger\OneExcel\Reader\SpoutReader;

class OneExcelReaderFactory
{
    private $driver = Driver::AUTO;
    private $input_format;
    private $input_filename;

    /**
     * @return OneExcelReaderFactory
     */
    public static function create()
    {
        return new OneExcelReaderFactory();
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
     * @return OneExcelReader
     */
    public function make() {
        if (!empty($this->input_filename)) {
            $this->autoDetectFormatFromFilename($this->input_format, $this->input_filename);
        }

        if ($this->driver != Driver::AUTO) {
            $driver = $this->getDriverByName($this->driver);
        } else {
            $driver = $this->getDriverByFormat($this->input_format);
        }

        /** @var OneExcelReader $driver_impl */
        $driver_impl = new $driver;

        $driver_impl->load($this->input_filename, $this->input_format);

        return $driver_impl;
    }

    /**
     * @param $input_format
     * @param $filename
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
                return PHPExcelReader::class;
            case Driver::LIBXL:
                return LibXLReader::class;
            case Driver::SPOUT:
                return SpoutReader::class;
            case Driver::FPUTCSV:
                return FCsvReader::class;
        }
        throw new \Exception("Unknown driver {$driver}");
    }

    /**
     * @param $output_format
     * @param null $input_format
     * @return string
     */
    private function getDriverByFormat($output_format)
    {
        return PHPExcelReader::class;
    }
}