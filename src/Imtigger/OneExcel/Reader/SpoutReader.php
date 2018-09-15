<?php
namespace Imtigger\OneExcel\Reader;

use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelReaderInterface;

class SpoutReader extends OneExcelReader implements OneExcelReaderInterface
{
    public static $input_format_supported = [Format::XLSX, Format::CSV, Format::ODS];
    /** @var \Box\Spout\Reader\AbstractReader $reader */
    private $reader;
    /** @var \Box\Spout\Reader\SheetInterface $sheet */
    private $sheet;

    public function load($filename, $input_format = Format::AUTO, $options = [])
    {
        $this->checkFormatSupported($input_format);
        $this->input_format = $input_format;

        $this->reader = \Box\Spout\Reader\ReaderFactory::create($input_format);
        $this->reader->setShouldFormatDates(true);
        $this->reader->setShouldPreserveEmptyRows(true);
        $this->reader->open($filename);

        foreach ($this->reader->getSheetIterator() as $sheetIndex => $sheet) {
            $this->sheet = $sheet;
            break;
        }
    }

    public function row()
    {
        foreach ($this->sheet->getRowIterator() as $row) {
            yield $row;
        }
    }

    public function close()
    {
        $this->reader->close();
    }
}