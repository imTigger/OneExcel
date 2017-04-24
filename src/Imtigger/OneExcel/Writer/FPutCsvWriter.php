<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelWriterInterface;

class FPutCsvWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [Format::CSV];
    public static $output_format_supported = [Format::CSV];
    public static $input_output_same_format = true;
    private $input_format;
    private $output_format;
    private $last_row = 0;
    private $data = [];
    private $temp_file;
    private $handle;

    public function create($output_format = Format::CSV)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;
        $this->temp_file = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'fputcsv-' . time();
        $this->handle = fopen($this->temp_file, 'w');
    }

    public function load($filename, $output_format = Format::XLSX, $input_format = Format::AUTO)
    {
        $this->checkFormatSupported($output_format, $input_format);

        $this->input_format = $input_format;
        $this->output_format = $output_format;

        if (($handle = fopen($filename, "r")) !== FALSE) {
            $this->create($output_format);
            while (($row = fgetcsv($handle)) !== FALSE) {
                // Skip empty rows
                if(!array_filter($row)) {
                    continue;
                }
                fputcsv($this->handle, $row);
                $this->last_row += 1;
            }
            fclose($handle);
        } else {
            throw new \Exception("{$filename} cannot be opened.");
        }
    }

    public function writeCell($row_num, $column_num, $data, $data_type = null)
    {
        if ($row_num < $this->last_row) {
            throw new \Exception("Writing row backward is not supported by fputcsv, was on row {$this->last_row}, trying to write row {$row_num}");
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

    private function close()
    {
        $this->flushRow();
        fclose($this->handle);
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

        fputcsv($this->handle, $data);
    }

    private function flushRow() {
        if (!empty($this->data)) {
            $this->addRow($this->data);
            $this->last_row += 1;
            $this->data = [];
        }
    }
}