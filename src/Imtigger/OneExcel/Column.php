<?php
namespace Imtigger\OneExcel;
use PHPExcel_Cell;

class Column
{
    public static function index($name) {
        return PHPExcel_Cell::columnIndexFromString($name) - 1;
    }

    public static function name($index) {
        return PHPExcel_Cell::stringFromColumnIndex($index - 1);
    }
}