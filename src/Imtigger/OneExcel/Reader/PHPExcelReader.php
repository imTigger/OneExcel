<?php
namespace Imtigger\OneExcel\Reader;

use Imtigger\OneExcel\ColumnType;
use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelReaderInterface;
use PHPExcel_Cell_DataType;
use PHPExcel_IOFactory;

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

        $objReader = PHPExcel_IOFactory::createReader($this->getFormatCode($this->input_format));

        $this->book = $objReader->load($filename);
        $this->sheet = $this->book->getActiveSheet();
    }

    public function row()
    {
        $rowIterator = $this->sheet->getRowIterator();

        foreach ($rowIterator As $row) {
            $cellIterator = $row->getCellIterator();
            $data = [];
            foreach ($cellIterator As $cell) {
                $data[] = $cell->getValue();
            }
            yield $data;
        }
    }

    public function close()
    {
        unset($this->sheet);
        unset($this->book);
    }
}