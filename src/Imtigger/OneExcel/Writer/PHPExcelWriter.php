<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\OneExcelWriterInterface;
use PHPExcel_Cell_DataType;
use PHPExcel_IOFactory;

class PHPExcelWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [self::FORMAT_XLSX, self::FORMAT_XLS, self::FORMAT_CSV, self::FORMAT_ODS];
    public static $output_format_supported = [self::FORMAT_XLSX, self::FORMAT_XLS, self::FORMAT_CSV];
    public static $input_output_same_format = false;
    /** @var \PHPExcel $book */
    private $book;
    /** @var \PHPExcel_Worksheet $sheet */
    private $sheet;
    private $input_format;
    private $output_format;

    public function create($output_format = self::FORMAT_XLSX)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;
        $this->book = new \PHPExcel();
        $this->sheet = $this->book->getActiveSheet();
    }

    public function load($filename, $output_format = self::FORMAT_XLSX, $input_format = self::FORMAT_AUTO)
    {
        $this->checkFormatSupported($output_format, $input_format);

        $this->input_format = $input_format;
        $this->output_format = $output_format;

        $objReader = PHPExcel_IOFactory::createReader($this->getFormatCode($this->input_format));

        $this->book = $objReader->load($filename);
        $this->sheet = $this->book->getActiveSheet();
    }

    public function writeCell($row_num, $column_num, $data, $data_type = self::COLUMN_TYPE_STRING)
    {
        $this->sheet->setCellValueExplicitByColumnAndRow($column_num, $row_num, $data, $this->getColumnFormat($data_type));
    }

    public function download($filename)
    {
        header('Content-Type: ' . $this->getFormatMime($this->output_format));
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Pragma: no-cache');

        $this->save('php://output');
        exit;
    }

    public function save($path)
    {
        /** @var \PHPExcel_Writer_Abstract $objWriter */
        $objWriter = PHPExcel_IOFactory::createWriter($this->book, $this->getFormatCode($this->output_format));
        $objWriter->setPreCalculateFormulas(false);
        $objWriter->save($path);
    }

    public function getFormatCode($format)
    {
        switch ($format) {
            case self::FORMAT_XLSX:
                return 'Excel2007';
            case self::FORMAT_XLS:
                return 'Excel5';
            case self::FORMAT_CSV:
                return 'CSV';
            case self::FORMAT_ODS:
                return 'OOCalc';
        }
        throw new \Exception("Unknown format {$format}");
    }

    public function getColumnFormat($internal_format)
    {
        switch ($internal_format) {
            case self::COLUMN_TYPE_STRING:
                return PHPExcel_Cell_DataType::TYPE_STRING;
            case self::COLUMN_TYPE_NUMERIC:
                return PHPExcel_Cell_DataType::TYPE_NUMERIC;
            case self::COLUMN_TYPE_BOOLEAN:
                return PHPExcel_Cell_DataType::TYPE_BOOL;
            case self::COLUMN_TYPE_FORMULA:
                return PHPExcel_Cell_DataType::TYPE_FORMULA;
            case self::COLUMN_TYPE_NULL:
                return PHPExcel_Cell_DataType::TYPE_NULL;
        }
        return PHPExcel_Cell_DataType::TYPE_STRING;
    }
}