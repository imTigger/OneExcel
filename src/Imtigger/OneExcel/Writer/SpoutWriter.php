<?php
namespace Imtigger\OneExcel\Writer;

use Imtigger\OneExcel\OneExcelWriterInterface;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Common\Type;

class SpoutWriter extends OneExcelWriter implements OneExcelWriterInterface
{
    public static $format_supported = [self::FORMAT_XLSX, self::FORMAT_CSV, self::FORMAT_ODS];
    private $writer;
    private $format;
    private $last_row = -1;
    private $data = [];
    private $temp_file;

    public function create($format = self::FORMAT_XLSX)
    {
        $this->checkFormatSupported($format);
        $this->format = $format;
        $this->temp_file = sys_get_temp_dir() . 'spout-' . time();
        $this->writer = WriterFactory::create($format);
        $this->writer->openToFile($this->temp_file);
    }

    public function load($filename, $format = self::FORMAT_XLSX)
    {
        $this->checkFormatSupported($format);
        $this->format = $format;
        throw new \Exception('SpoutWriter::load is not implemented');
    }

    public function writeCell($row_num, $column_num, $data, $data_type = null)
    {
        if ($row_num < $this->last_row) {
            throw new \Exception('Row rewind is not supported in Spout');
        }

        // Aggregate columns in same row
        if ($this->last_row == $row_num) {
            $this->data[$column_num] = $data;
        } else {
            $this->writer->addRow($this->data);
            $this->data = [];
        }
        $this->last_row = $row_num;
    }

    public function close()
    {
        // Write out last row
        $this->writer->addRow($this->data);
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

        header('Content-Type: ' . $this->getFormatMime($this->format));
        header('Content-Disposition: attachment; filename="spout-' . $filename . '.' . $this->format . '"');
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Pragma: no-cache');

        echo file_get_contents($this->temp_file);

        @unlink($this->temp_file);

        exit;
    }
}