<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\OneExcelWriterInterface;

class LibXLWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $format_supported = [self::FORMAT_XLSX, self::FORMAT_XLS];
    private $book;
    private $sheet;
    private $format;

    public function create($format = self::FORMAT_XLSX)
    {
        $this->checkFormatSupported($format);
        $this->format = $format;
        $this->book = new \ExcelBook(null, null, $this->format == self::FORMAT_XLSX);
        $this->book->setLocale('UTF-8');
        $this->sheet = $this->book->addSheet('Sheet1');
    }

    public function load($filename, $format = self::FORMAT_XLSX)
    {
        $this->checkFormatSupported($format);
        $this->format = $format;
        $this->book = new \ExcelBook(null, null, $this->format == self::FORMAT_XLSX);
        $this->book->loadFile($filename);
        $this->book->setLocale('UTF-8');
        $this->sheet = $this->book->getSheet(0);
    }

    public function writeColumn($row_num, $column_num, $data, $data_type = null)
    {
        $this->sheet->write($row_num - 1, $column_num, $data);
    }

    public function save($path)
    {
        $this->book->save($path);
    }

    public function download($filename)
    {
        header('Content-Type: ' . $this->getFormatMime($this->format));
        header('Content-Disposition: attachment; filename="libxl-' . $filename . '.' . $this->format . '"');
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Pragma: no-cache');

        $this->save('php://output');
        exit;
    }
}