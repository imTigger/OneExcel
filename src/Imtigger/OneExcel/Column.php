<?php
namespace Imtigger\OneExcel;

class Column
{
    public static function index($name) {
        return \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($name) - 1;
    }

    public static function name($index) {
        return \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($index - 1);
    }
}