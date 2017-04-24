<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelWriterInterface;

class FPutCsvWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [Format::CSV];
    public static $output_format_supported = [Format::CSV];
    public static $input_output_same_format = true;
    private $last_row = 0;
    private $data = [];
    private $handle;
    private $temp_file;

    public function create($output_format = Format::CSV)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;

        if ($this->output_mode == 'stream') {
            header('Content-Type: ' . $this->getFormatMime($this->output_format));
            header('Content-Disposition: attachment; filename="' . $this->output_filename . '"');
            header('Content-Transfer-Encoding: binary');
            header('Expires: 0');
            header('Pragma: no-cache');

            $this->handle = fopen("php://output", 'w');
        } elseif ($this->output_mode == 'download') {
            $this->temp_file = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'fputcsv-' . time() . '.tmp';
            $this->handle = fopen($this->temp_file, 'w');
        } elseif ($this->output_mode == 'file') {
            $this->handle = fopen($this->output_filename, 'w');
        }
    }

    public function load($filename, $output_format = Format::CSV, $input_format = Format::CSV)
    {
        $this->checkFormatSupported($output_format, $input_format);
        $this->input_format = $input_format;

        $this->create($output_format);

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

    public function output()
    {
        $this->close();

        if ($this->output_mode == 'download') {
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

    public function download()
    {
        $this->output_mode = 'download';
        $this->output();
    }

    /* Private helpers */
    private function close()
    {
        $this->flushRow();
        fclose($this->handle);
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