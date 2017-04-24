<?php
namespace Imtigger\OneExcel;


interface OneExcelWriterInterface
{
    public function create($output_format = Format::XLSX);

    public function load($filename, $output_format = Format::XLSX, $input_format = Format::AUTO);

    public function writeCell($row_num, $column_num, $data, $data_type = ColumnType::STRING);

    public function writeRow($row_num, $data);

    public function output();
}