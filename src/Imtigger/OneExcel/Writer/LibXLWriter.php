<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\ColumnType;
use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelWriterInterface;

class LibXLWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [Format::XLSX, Format::XLS];
    public static $output_format_supported = [Format::XLSX, Format::XLS];
    public static $input_output_same_format = true;
    /** @var \ExcelBook $book */
    private $book;
    /** @var \ExcelSheet $sheet */
    private $sheet;

    public function create($output_format = Format::XLSX)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;
        $this->book = new \ExcelBook(null, null, $this->output_format == Format::XLSX);
        $this->book->setLocale('UTF-8');
        $this->sheet = $this->book->addSheet('Sheet1');
    }

    public function load($filename, $output_format = Format::XLSX, $input_format = Format::AUTO)
    {
        $this->checkFormatSupported($output_format, $input_format);

        $this->input_format = $input_format;
        $this->output_format = $output_format;

        $this->book = new \ExcelBook(null, null, $this->output_format == Format::XLSX);
        $this->book->loadFile($filename);
        $this->book->setLocale('UTF-8');
        $this->sheet = $this->book->getSheet(0);
    }

    public function writeCell($row_num, $column_num, $data, $data_type = null)
    {
        $this->sheet->write($row_num - 1, $column_num, $data, $this->getCellFormat($data_type), $this->getCellDataType($data_type));
    }

    public function writeRow($row_num, $data)
    {
        $this->sheet->writeRow($row_num - 1, $data);
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
        $this->book->save($path);
    }

    private function getCellFormat($internal_format)
    {
        $format = new \ExcelFormat($this->book);

        switch ($internal_format) {
            case ColumnType::STRING:
                $format->numberFormat(\ExcelFormat::NUMFORMAT_TEXT);
                return $format;
            case ColumnType::TIME:
                $format->numberFormat(\ExcelFormat::NUMFORMAT_CUSTOM_HMMSS);
                return $format;
            case ColumnType::DATE:
                $newFormat = $this->book->addCustomFormat('yyyy-mm-dd');
                $format->numberFormat($newFormat);
                return $format;
            case ColumnType::DATETIME:
                $newFormat = $this->book->addCustomFormat('yyyy-mm-dd hh:mm:ss');
                $format->numberFormat($newFormat);
                return $format;
            case ColumnType::INTEGER:
                $format->numberFormat(\ExcelFormat::NUMFORMAT_NUMBER);
                return $format;
            case ColumnType::NUMERIC:
                $format->numberFormat(\ExcelFormat::NUMFORMAT_NUMBER_D2);
                return $format;
            case ColumnType::FORMULA:
                return null;
        }
        return null;
    }

    private function getCellDataType($internal_format)
    {
        switch ($internal_format) {
            case ColumnType::STRING:
                return -1;
            case ColumnType::DATE:
            case ColumnType::DATETIME:
                return \ExcelFormat::AS_DATE;
            case ColumnType::NUMERIC:
                return \ExcelFormat::AS_NUMERIC_STRING;
            case ColumnType::FORMULA:
                return \ExcelFormat::AS_FORMULA;
        }
        return null;
    }
}