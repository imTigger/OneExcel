<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelWriterInterface;

class SpoutWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [Format::XLSX, Format::CSV, Format::ODS];
    public static $output_format_supported = [Format::XLSX, Format::CSV, Format::ODS];
    public static $input_output_same_format = false;
    /** @var \Box\Spout\Writer\AbstractMultiSheetsWriter $writer */
    private $writer;
    private $last_row = 0;
    private $data = [];
    private $temp_file;

    public function create($output_format = Format::XLSX)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;

        $this->writer = \Box\Spout\Writer\WriterFactory::create($this->output_format);
        if ($this->output_mode == OneExcelWriter::OUTPUT_MODE_STREAM) {
            $this->writer->openToBrowser($this->output_filename);
        } elseif ($this->output_mode == OneExcelWriter::OUTPUT_MODE_DOWNLOAD || $this->output_mode == null) {
            $this->temp_file = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'spout-' . time() . '.tmp';
            $this->writer->openToFile($this->temp_file);
        } elseif ($this->output_mode == OneExcelWriter::OUTPUT_MODE_FILE) {
            $this->writer->openToFile($this->output_filename);
        }
    }

    public function load($filename, $output_format = Format::XLSX, $input_format = Format::AUTO, $options = [])
    {
        $this->checkFormatSupported($output_format, $input_format);
        $this->input_format = $input_format;

        $this->create($output_format);

        // Copy data into new sheet
        /** @var \Box\Spout\Reader\AbstractReader $reader */
        $reader = \Box\Spout\Reader\ReaderFactory::create($input_format);
        $reader->open($filename);
        $reader->setShouldFormatDates(true);

        foreach ($reader->getSheetIterator() as $sheetIndex => $sheet) {
            if ($sheetIndex !== 1) {
                $this->writer->addNewSheetAndMakeItCurrent();
            }

            foreach ($sheet->getRowIterator() as $row) {
                // Skip empty rows
                if(!array_filter($row)) {
                    continue;
                }
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

    public function output()
    {
        $this->close();

        if ($this->output_mode == OneExcelWriter::OUTPUT_MODE_DOWNLOAD) {
            header('Content-Type: ' . $this->getFormatMime($this->output_format));
            header('Content-Disposition: attachment; filename="' . $this->output_filename . '"');
            header('Content-Transfer-Encoding: binary');
            header('Expires: 0');
            header('Pragma: no-cache');
            header('Content-Length: ' . filesize($this->temp_file));

            echo file_get_contents($this->temp_file);
            //unlink($this->temp_file);
        }
    }

    /* Private helpers */
    private function close()
    {
        $this->flushRow();
        $this->writer->close();
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