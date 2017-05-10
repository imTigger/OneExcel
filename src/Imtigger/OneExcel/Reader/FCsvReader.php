<?php
namespace Imtigger\OneExcel\Reader;

use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelReaderInterface;

class FCsvReader extends OneExcelReader implements OneExcelReaderInterface
{
    public static $input_format_supported = [Format::CSV];
    private $handle;

    public function load($filename, $output_format = Format::CSV, $input_format = Format::CSV)
    {
        $this->checkFormatSupported($input_format);
        $this->input_format = $input_format;

        $this->handle = fopen($filename, "r");

        if ($this->handle === FALSE) {
            throw new \Exception("{$filename} cannot be opened.");
        }
    }


    public function row()
    {
        while (($row = fgetcsv($this->handle)) !== FALSE) {
            yield $row;
        }
    }

    public function close()
    {
        fclose($this->handle);
    }
}