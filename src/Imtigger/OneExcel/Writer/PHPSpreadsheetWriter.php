<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\ColumnType;
use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelWriterInterface;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\BaseWriter;

class PHPSpreadsheetWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [Format::XLSX, Format::XLS, Format::CSV, Format::ODS];
    public static $output_format_supported = [Format::XLSX, Format::XLS, Format::CSV];
    public static $input_output_same_format = false;
    /** @var Spreadsheet $book */
    private $book;
    /** @var Worksheet $sheet */
    private $sheet;

    public function create($output_format = Format::XLSX)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;

        $this->book = new Spreadsheet;
        $this->sheet = $this->book->getActiveSheet();
    }

    public function load($filename, $output_format = Format::XLSX, $input_format = Format::AUTO)
    {
        $this->checkFormatSupported($output_format, $input_format);

        $this->input_format = $input_format;
        $this->output_format = $output_format;

        $objReader = IOFactory::createReader($this->getFormatCode($this->input_format));

        $this->book = $objReader->load($filename);
        $this->sheet = $this->book->getActiveSheet();
    }

    public function writeCell($row_num, $column_num, $data, $data_type = null)
    {
        $cell = $this->sheet->getCellByColumnAndRow($column_num + 1, $row_num);
        $cell->setValueExplicit($data, $this->getCellDataType($data_type));
        if ($data_type != null) {
            $cell->getStyle()->getNumberFormat()->setFormatCode($this->getCellFormat($data_type));
        }
    }

    public function writeRow($row_num, $data)
    {
        $this->sheet->fromArray($data, null, "A{$row_num}");
    }

    public function output()
    {
        if ($this->output_mode == 'stream' || $this->output_mode == 'download') {
            header('Content-Type: ' . $this->getFormatMime($this->output_format));
            header('Content-Disposition: attachment; filename="' . $this->output_filename . '"');
            header('Content-Transfer-Encoding: binary');
            header('Expires: 0');
            header('Pragma: no-cache');

            $this->saveFile('php://output');
        } elseif ($this->output_mode == 'file') {
            $this->saveFile($this->output_filename);
        }
    }

    /* Private helpers */
    private function saveFile($path)
    {
        /** @var BaseWriter $objWriter */
        $objWriter = IOFactory::createWriter($this->book, $this->getFormatCode($this->output_format));
        $objWriter->setPreCalculateFormulas(false);
        $objWriter->save($path);
    }

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

    private function getCellFormat($internal_format)
    {
        switch ($internal_format) {
            case ColumnType::STRING:
                return NumberFormat::FORMAT_TEXT;
            case ColumnType::TIME:
                return 'hh:mm:ss';
            case ColumnType::DATE:
                return 'yyyy-mm-dd';
            case ColumnType::DATETIME:
                return 'yyyy-mm-dd hh:mm:ss';
            case ColumnType::INTEGER:
                return NumberFormat::FORMAT_NUMBER;
            case ColumnType::NUMERIC:
                return NumberFormat::FORMAT_NUMBER_00;
            case ColumnType::FORMULA:
                return null;
        }
        return null;
    }

    private function getCellDataType($internal_format)
    {
        switch ($internal_format) {
            case ColumnType::STRING:
                return DataType::TYPE_STRING;
            case ColumnType::INTEGER:
            case ColumnType::NUMERIC:
                return DataType::TYPE_NUMERIC;
            case ColumnType::BOOLEAN:
                return DataType::TYPE_BOOL;
            case ColumnType::FORMULA:
                return DataType::TYPE_FORMULA;
            case ColumnType::NULL:
                return DataType::TYPE_NULL;
        }
        return DataType::TYPE_STRING;
    }
}