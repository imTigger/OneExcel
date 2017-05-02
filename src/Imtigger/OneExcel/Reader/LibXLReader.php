<?php
namespace Imtigger\OneExcel\Reader;

use Imtigger\OneExcel\ColumnType;
use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelReaderInterface;
use Imtigger\OneExcel\OneExcelWriterInterface;
use Imtigger\OneExcel\Reader\OneExcelReader;

class LibXLReader extends OneExcelReader implements OneExcelReaderInterface
{
    public static $input_format_supported = [Format::XLSX, Format::XLS];
    /** @var \ExcelBook $book */
    private $book;
    /** @var \ExcelSheet $sheet */
    private $sheet;

    public function load($filename, $input_format = Format::AUTO)
    {
        $this->checkFormatSupported($input_format);

        $this->input_format = $input_format;

        $this->book = new \ExcelBook(null, null, $this->input_format == Format::XLSX);
        $this->book->loadFile($filename);
        $this->book->setLocale('UTF-8');
        $this->sheet = $this->book->getSheet(0);
    }

    public function row()
    {
        for ($i = $this->sheet->firstRow(); $i < $this->sheet->lastRow(); $i += 1) {
            $data = [];
            for($c = $this->sheet->firstCol(); $c < $this->sheet->lastCol(); $c += 1) {
                $data[] = $this->sheet->read($i, $c);
            }
            yield $data;
        }
    }

    public function close()
    {
        unset($this->book);
    }
}