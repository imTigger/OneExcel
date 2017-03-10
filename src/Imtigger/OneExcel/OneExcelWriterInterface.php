<?php
namespace Imtigger\OneExcel;


interface OneExcelWriterInterface
{
    const COLUMN_TYPE_STRING = 0;
    const COLUMN_TYPE_NUMERIC = 1;

    const FORMAT_XLSX = 'xlsx';
    const FORMAT_XLS = 'xls';
    const FORMAT_CSV = 'csv';
    const FORMAT_ODS = 'ods';

    public function create($format = self::FORMAT_XLSX);
    public function load($filename, $format = self::FORMAT_XLSX);
    public function writeColumn($row_num, $column_num, $data, $data_type = self::COLUMN_TYPE_STRING);
    public function save($path);
    public function download($filename);
}