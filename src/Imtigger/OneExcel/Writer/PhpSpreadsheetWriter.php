<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\ColumnType;
use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelWriterInterface;

class PhpSpreadsheetWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [Format::XLSX, Format::XLS, Format::CSV, Format::ODS];
    public static $output_format_supported = [Format::XLSX, Format::XLS, Format::CSV];
    public static $input_output_same_format = false;
    /** @var \PhpOffice\PhpSpreadsheet\Spreadsheet $book */
    private $book;
    /** @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet */
    private $sheet;

    public function create($output_format = Format::XLSX)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;

        $this->book = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $this->sheet = $this->book->getActiveSheet();
    }

    public function load($filename, $output_format = Format::XLSX, $input_format = Format::AUTO)
    {
        $this->checkFormatSupported($output_format, $input_format);

        $this->input_format = $input_format;
        $this->output_format = $output_format;

        $objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($this->getFormatCode($this->input_format));

        $this->book = $objReader->load($filename);
        $this->sheet = $this->book->getActiveSheet();
    }
    
    public function setSheet($id)
    {
        if ($this->output_format == Format::XLSX || $this->output_format == Format::XLS) {
            $this->sheet = $this->book->setActiveSheetIndex($id);
        } else {
            throw new \Exception("Unsupported format {$this->output_format}");
        }
    }

    public function writeCell($row_num, $column_num, $data, $data_type = null)
    {
        // PhpSpreadSheet: Column indexes are now based on 1, row indexes also based on 1
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
        if ($this->output_mode == OneExcelWriter::OUTPUT_MODE_STREAM || $this->output_mode == OneExcelWriter::OUTPUT_MODE_DOWNLOAD) {
            header('Content-Type: ' . $this->getFormatMime($this->output_format));
            header('Content-Disposition: attachment; filename="' . $this->output_filename . '"');
            header('Content-Transfer-Encoding: binary');
            header('Expires: 0');
            header('Pragma: no-cache');

            $this->saveFile('php://output');
        } elseif ($this->output_mode == OneExcelWriter::OUTPUT_MODE_FILE) {
            $this->saveFile($this->output_filename);
        }
    }

    /* Private helpers */
    private function saveFile($path)
    {
        /** @var \PhpOffice\PhpSpreadsheet\Writer\BaseWriter $objWriter */
        $objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->book, $this->getFormatCode($this->output_format));
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
                return \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT;
            case ColumnType::TIME:
                return 'hh:mm:ss';
            case ColumnType::DATE:
                return 'yyyy-mm-dd';
            case ColumnType::DATETIME:
                return 'yyyy-mm-dd hh:mm:ss';
            case ColumnType::INTEGER:
                return \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER;
            case ColumnType::NUMERIC:
                return \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER_00;
            case ColumnType::FORMULA:
                return null;
        }
        return null;
    }

    private function getCellDataType($internal_format)
    {
        switch ($internal_format) {
            case ColumnType::STRING:
                return \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING;
            case ColumnType::INTEGER:
            case ColumnType::NUMERIC:
                return \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC;
            case ColumnType::BOOLEAN:
                return \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_BOOL;
            case ColumnType::FORMULA:
                return \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA;
            case ColumnType::NULL:
                return \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL;
        }
        return \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING;
    }
}
