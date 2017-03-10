<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\OneExcelWriterInterface;
use PHPExcel_Cell_DataType;
use PHPExcel_IOFactory;

class PHPExcelWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $format_supported = [self::FORMAT_XLSX, self::FORMAT_XLS, self::FORMAT_CSV, self::FORMAT_ODS];
    private $book;
    private $sheet;
    private $format;

    public function create($format = self::FORMAT_XLSX)
    {
        $this->checkFormatSupported($format);
        $this->format = $format;
        $this->book = new \PHPExcel();
        $this->sheet = $this->book->getActiveSheet();
    }

    public function load($filename, $format = self::FORMAT_XLSX)
    {
        $this->checkFormatSupported($format);
        $this->format = $format;
        $objReader = PHPExcel_IOFactory::createReader($this->getFormatCode($this->format));
        $this->book = $objReader->load($filename);
        $this->sheet = $this->book->getActiveSheet();
    }

    public function writeColumn($row_num, $column_num, $data, $data_type = null)
    {
        $this->sheet->setCellValueExplicitByColumnAndRow($column_num, $row_num, $data, PHPExcel_Cell_DataType::TYPE_STRING);
    }

    public function save($path)
    {
        $objWriter = PHPExcel_IOFactory::createWriter($this->book, $this->getFormatCode($this->format));
        $objWriter->setPreCalculateFormulas(false);
        $objWriter->save($path);
    }

    public function download($filename)
    {
        header('Content-Type: ' . $this->getFormatMime($this->format));
        header('Content-Disposition: attachment; filename="phpexcel-' . $filename . '.' . $this->format . '"');
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Pragma: no-cache');

        $this->save('php://output');
        exit;
    }

    public function getFormatCode($format) {
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
    }
}