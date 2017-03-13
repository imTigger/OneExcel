<?php
namespace Imtigger\OneExcel;


interface OneExcelWriterInterface
{
    const COLUMN_TYPE_STRING = 0;
    const COLUMN_TYPE_NUMERIC = 1;
    const COLUMN_TYPE_BOOLEAN = 2;
    const COLUMN_TYPE_DATE = 3;
    const COLUMN_TYPE_TIME = 4;
    const COLUMN_TYPE_DATETIME = 5;
    const COLUMN_TYPE_FORMULA = 6;
    const COLUMN_TYPE_NULL = 7;

    const FORMAT_AUTO = 'AUTO';
    const FORMAT_XLSX = 'xlsx';
    const FORMAT_XLS = 'xls';
    const FORMAT_CSV = 'csv';
    const FORMAT_ODS = 'ods';

    public function create($output_format = self::FORMAT_XLSX);

    public function load($filename, $output_format = self::FORMAT_XLSX, $input_format = self::FORMAT_AUTO);

    public function writeCell($row_num, $column_num, $data, $data_type = self::COLUMN_TYPE_STRING);

    public function save($path);

    public function download($filename);
}