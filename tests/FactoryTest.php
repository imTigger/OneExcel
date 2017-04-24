<?php

use Imtigger\OneExcel\Driver;
use Imtigger\OneExcel\Format;
use PHPUnit\Framework\TestCase;

final class FactoryTest extends TestCase
{
    private function requireLibXL()
    {
        if (!extension_loaded('excel')) {
            $this->markTestSkipped(
                'The LibXL extension is not available.'
            );
        }
    }

    public function testCreate()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create(Format::XLSX);
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreatePHPExcel()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create(Format::XLSX, Driver::PHPEXCEL);
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\PHPExcelWriter::class, $excel);
    }

    public function testCreateLibXL()
    {
        $this->requireLibXL();

        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create(Format::XLSX, Driver::LIBXL);
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\LibXLWriter::class, $excel);
    }

    public function testCreateSpout()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create(Format::XLSX, Driver::SPOUT);
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\SpoutWriter::class, $excel);
    }

    public function testCreateFputcsv()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create(Format::CSV, Driver::FPUTCSV);
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\FPutCsvWriter::class, $excel);
    }

    public function testUnsupportedDriver()
    {
        $this->expectException(\Exception::class);
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create(Format::XLSX, Driver::FPUTCSV);
    }

    public function testCreateFromFileXLSX()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::createFromFile(__DIR__ . '/../templates/template.xlsx');
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileXLS()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::createFromFile(__DIR__ . '/../templates/template.xls');
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileCSV()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::createFromFile(__DIR__ . '/../templates/template.csv');
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileODS()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::createFromFile(__DIR__ . '/../templates/template.ods');
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileConvert()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::createFromFile(__DIR__ . '/../templates/template.xlsx', Format::CSV);
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileConvertWithDriver()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::createFromFile(__DIR__ . '/../templates/template.xlsx', Format::CSV, Format::XLSX, Driver::SPOUT);
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileConvertWithDriverUnsupportedInputOutput()
    {
        $this->expectException(\Exception::class);

        $this->requireLibXL();

        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::createFromFile(__DIR__ . '/../templates/template.xlsx', Format::XLS, Format::XLSX, Driver::LIBXL);
    }

    public function testCreateFromFileConvertWithDriverUnsupportedInput()
    {
        $this->expectException(\Exception::class);

        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::createFromFile(__DIR__ . '/../templates/template.csv', Format::XLSX, Format::CSV, Driver::FPUTCSV);
    }

    public function testCreateFromFileConvertWithDriverUnsupportedOutput()
    {
        $this->expectException(\Exception::class);

        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::createFromFile(__DIR__ . '/../templates/template.xlsx', Format::CSV, Format::XLSX, Driver::FPUTCSV);
    }
}
