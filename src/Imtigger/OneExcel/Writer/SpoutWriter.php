<?php
namespace Imtigger\OneExcel\Writer;

use Box\Spout\Writer\AbstractMultiSheetsWriter;
use Imtigger\OneExcel\OneExcelWriterInterface;
use Box\Spout\Writer\WriterFactory;

class SpoutWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $input_format_supported = [self::FORMAT_XLSX, self::FORMAT_CSV, self::FORMAT_ODS];
    public static $output_format_supported = [self::FORMAT_XLSX, self::FORMAT_CSV, self::FORMAT_ODS];
    public static $input_output_same_format = false;
    /** @var AbstractMultiSheetsWriter $writer */
    private $writer;
    private $input_format;
    private $output_format;
    private $last_row = -1;
    private $data = [];
    private $temp_file;

    public function create($output_format = self::FORMAT_XLSX)
    {
        $this->checkFormatSupported($output_format);
        $this->output_format = $output_format;
        $this->temp_file = sys_get_temp_dir() . 'spout-' . time();
        $this->writer = WriterFactory::create($output_format);
        $this->writer->openToFile($this->temp_file);
    }

    public function load($filename, $output_format = self::FORMAT_XLSX, $input_format = self::FORMAT_AUTO)
    {
        $this->checkFormatSupported($output_format, $input_format);

        $this->input_format = $input_format;
        $this->output_format = $output_format;

        throw new \Exception('SpoutWriter::load is not implemented');
    }

    public function writeCell($row_num, $column_num, $data, $data_type = null)
    {
        if ($row_num < $this->last_row) {
            throw new \Exception('Row rewind is not supported in Spout');
        }

        // Aggregate columns in same row
        if ($this->last_row != $row_num) {
            $this->addRow($this->data);
            $this->data = [];
        }

        $this->data[$column_num] = $data;

        $this->last_row = $row_num;
    }

    public function close()
    {
        // Write out last row
        $this->addRow($this->data);
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
}