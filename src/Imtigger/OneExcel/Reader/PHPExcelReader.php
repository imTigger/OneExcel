<?php
namespace Imtigger\OneExcel\Reader;

use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelReaderInterface;
use PHPExcel_IOFactory;
use PHPExcel_CachedObjectStorageFactory;
use PHPExcel_Settings;

class PHPExcelReader extends OneExcelReader implements OneExcelReaderInterface
{
    public static $input_format_supported = [Format::XLSX, Format::XLS, Format::CSV, Format::ODS];
    /** @var \PHPExcel $book */
    private $book;
    /** @var \PHPExcel_Worksheet $sheet */
    private $sheet;

    public function load($filename, $input_format = Format::AUTO)
    {
        $this->checkFormatSupported($input_format);

        $this->input_format = $input_format;

        $cacheMethod = PHPExcel_CachedObjectStorageFactory:: cache_to_phpTemp;
        $cacheSettings = array( ' memoryCacheSize ' => '8MB');
        PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);
        
        $objReader = PHPExcel_IOFactory::createReader($this->getFormatCode($this->input_format));
        $objReader->setReadDataOnly(true);

        $this->book = $objReader->load($filename);
        $this->sheet = $this->book->getActiveSheet();
    }

    public function row()
    {
        $rowIterator = $this->sheet->getRowIterator();

        foreach ($rowIterator As $row) {
            $cellIterator = $row->getCellIterator();
            $data = [];
            /** @var \PHPExcel_Cell $cell */
            foreach ($cellIterator As $cell) {
                $value = $cell->getValue();

                if ($value instanceof \PHPExcel_RichText) {
                    $value = $value->getPlainText();
                }

                $data[] = $value;
            }
            yield $data;
        }
    }

    public function close()
    {
        unset($this->sheet);
        unset($this->book);
    }

    /* Private helpers */
    private function getFormatCode($format)
    {
        switch ($format) {
            case Format::XLSX:
                return 'Excel2007';
            case Format::XLS:
                return 'Excel5';
            case Format::CSV:
                return 'CSV';
            case Format::ODS:
                return 'OOCalc';
        }
        throw new \Exception("Unknown format {$format}");
    }
}
