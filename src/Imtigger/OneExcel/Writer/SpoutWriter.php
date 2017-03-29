<?php
namespace Imtigger\OneExcel\Writer;

use Box\Spout\Writer\AbstractMultiSheetsWriter;
use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelWriterInterface;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\WriterFactory;

class SpoutWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [Format::XLSX, Format::CSV, Format::ODS];
    public static $output_format_supported = [Format::XLSX, Format::CSV, Format::ODS];
    public static $input_output_same_format = false;
    /** @var AbstractMultiSheetsWriter $writer */
    private $writer;
    private $input_format;
    private $output_format;
    private $last_row = 0;
    private $data = [];
    private $temp_file;

    public function create($output_format = Format::XLSX)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;
        $this->temp_file = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'spout-' . time();
        $this->writer = WriterFactory::create($output_format);
        $this->writer->openToFile($this->temp_file);
    }

    public function load($filename, $output_format = Format::XLSX, $input_format = Format::AUTO, $options = [])
    {
        $this->checkFormatSupported($output_format, $input_format);

        $this->input_format = $input_format;
        $this->output_format = $output_format;

        $this->create($output_format);

        // Copy data into new sheet
        $reader = ReaderFactory::create($input_format);
        $reader->open($filename);
        $reader->setShouldFormatDates(true);

        foreach ($reader->getSheetIterator() as $sheetIndex => $sheet) {
            if ($sheetIndex !== 1) {
                $this->writer->addNewSheetAndMakeItCurrent();
            }

            foreach ($sheet->getRowIterator() as $row) {
                $this->writer->addRow($row);
                $this->last_row += 1;
            }
        }

        $reader->close();
    }

    public function writeCell($row_num, $column_num, $data, $data_type = null)
    {
        if ($row_num < $this->last_row) {
            throw new \Exception("Writing row backward is not supported by Spout, was on row {$this->last_row}, trying to write row {$row_num}");
        }

        // Aggregate columns in same row
        if ($this->last_row != $row_num) {
            $this->flushRow();
        }

        $this->data[$column_num] = $data;

        $this->last_row = $row_num;
    }

    public function writeRow($row_num, $data)
    {
        $this->flushRow();
        $this->addRow($data);
        $this->last_row += 1;
    }

    public function close()
    {
        $this->flushRow();
        $this->writer->close();
    }

    public function save($path)
    {
        $this->close();

        @copy($this->temp_file, $path);

        @unlink($this->temp_file);
    }

    public function download($filename)
    {
        $this->close();

        header('Content-Type: ' . $this->getFormatMime($this->output_format));
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Pragma: no-cache');

        echo file_get_contents($this->temp_file);

        @unlink($this->temp_file);

        exit;
    }

    private function addRow($data) {
        if (sizeof($data) == 0) return;

        // Pad empty cells
        for ($i = 0; $i <= max(array_keys($data)); $i += 1) {
            if (!isset($data[$i])) {
                $data[$i] = null;
            }
        }
        ksort($data);

        $this->writer->addRow($data);
    }

    private function flushRow() {
        if (!empty($this->data)) {
            $this->addRow($this->data);
            $this->last_row += 1;
            $this->data = [];
        }
    }
}