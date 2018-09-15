<?php
namespace Imtigger\OneExcel\Reader;

use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelReaderInterface;

class PhpSpreadsheetReader extends OneExcelReader implements OneExcelReaderInterface
{
    public static $input_format_supported = [Format::XLSX, Format::XLS, Format::CSV, Format::ODS];
    /** @var \PhpOffice\PhpSpreadsheet\Spreadsheet $book */
    private $book;
    /** @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet */
    private $sheet;

    public function load($filename, $input_format = Format::AUTO)
    {
        $this->checkFormatSupported($input_format);

        $this->input_format = $input_format;

        // TODO: Re-enable cache
        
        $objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($this->getFormatCode($this->input_format));
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
            /** @var \PhpOffice\PhpSpreadsheet\Cell\Cell $cell */
            foreach ($cellIterator As $cell) {
                $value = $cell->getValue();

                if ($value instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
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
                return 'Xlsx';
            case Format::XLS:
                return 'Xls';
            case Format::CSV:
                return 'Csv';
            case Format::ODS:
                return 'Ods';
        }
        throw new \Exception("Unknown format {$format}");
    }
}
