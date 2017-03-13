<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\OneExcelWriterInterface;

class LibXLWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [self::FORMAT_XLSX, self::FORMAT_XLS];
    public static $output_format_supported = [self::FORMAT_XLSX, self::FORMAT_XLS];
    public static $input_output_same_format = true;
    private $book;
    private $sheet;
    private $input_format;
    private $output_format;

    public function create($output_format = self::FORMAT_XLSX)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;
        $this->book = new \ExcelBook(null, null, $this->output_format == self::FORMAT_XLSX);
        $this->book->setLocale('UTF-8');
        $this->sheet = $this->book->addSheet('Sheet1');
    }

    public function load($filename, $output_format = self::FORMAT_XLSX, $input_format = self::FORMAT_AUTO)
    {
        $this->autoDetectInputFormat($filename, $input_format);
        $this->checkFormatSupported($output_format, $input_format);

        $this->input_format = $input_format;
        $this->output_format = $output_format;

        $this->book = new \ExcelBook(null, null, $this->output_format == self::FORMAT_XLSX);
        $this->book->loadFile($filename);
        $this->book->setLocale('UTF-8');
        $this->sheet = $this->book->getSheet(0);
    }

    public function writeCell($row_num, $column_num, $data, $data_type = self::COLUMN_TYPE_STRING)
    {
        $this->sheet->write($row_num - 1, $column_num, $data, null, $this->getColumnFormat($data_type));
    }

    public function save($path)
    {
        $this->book->save($path);
    }

    public function download($filename)
    {
        header('Content-Type: ' . $this->getFormatMime($this->output_format));
        header('Content-Disposition: attachment; filename="libxl-' . $filename . '.' . $this->output_format . '"');
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Pragma: no-cache');

        $this->save('php://output');
        exit;
    }

    public function getColumnFormat($internal_format)
    {
        switch ($internal_format) {
            case self::COLUMN_TYPE_STRING:
                return -1;
            case self::COLUMN_TYPE_NUMERIC:
                return ExcelFormat::AS_NUMERIC_STRING;
            case self::COLUMN_TYPE_FORMULA:
                return ExcelFormat::AS_FORMULA;
        }
        return -1;
    }
}