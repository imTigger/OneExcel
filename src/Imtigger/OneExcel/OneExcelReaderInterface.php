<?php
namespace Imtigger\OneExcel;


interface OneExcelReaderInterface
{
    public function load($filename, $input_format = Format::AUTO);
    
    public function setSheet($id);

    public function row();

    public function close();
}
