<?php

use Imtigger\OneExcel\Driver;
use Imtigger\OneExcel\Format;
use Imtigger\OneExcel\OneExcelWriterFactory;
use PHPUnit\Framework\TestCase;

final class ReaderFactoryTest extends TestCase
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
        $excel = OneExcelWriterFactory::create();
        $this->assertInstanceOf(\Imtigger\OneExcel\OneExcelWriterFactory::class, $excel);
    }

    public function testCreateXLSX()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile('test.xlsx')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateXLS()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile('test.xls')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateCSV()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile('test.csv')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateODS()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile('test.ods')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }


    public function testCreatePHPExcel()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile('test.xlsx')->withDriver(Driver::PHPEXCEL)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\PHPExcelWriter::class, $excel);
    }

    public function testCreateLibXL()
    {
        $this->requireLibXL();

        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile('test.xlsx')->withDriver(Driver::LIBXL)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\LibXLWriter::class, $excel);
    }

    public function testCreateSpout()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile('test.xlsx')->withDriver(Driver::SPOUT)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\SpoutWriter::class, $excel);
    }

    public function testCreateFputcsv()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile('test.csv')->withDriver(Driver::FPUTCSV)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\FCsvWriter::class, $excel);
    }

    public function testUnsupportedDriver()
    {
        $this->expectException(\Exception::class);
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile('test.xlsx')->withDriver(Driver::FPUTCSV)->make();
    }

    public function testCreateFromFileXLSX()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile(__DIR__ . '/../templates/template.xlsx')->toFile('test.xlsx')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileXLS()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile(__DIR__ . '/../templates/template.xls')->toFile('test.xlsx')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileCSV()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile(__DIR__ . '/../templates/template.csv')->toFile('test.csv')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileODS()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile(__DIR__ . '/../templates/template.ods')->toFile('test.ods')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileConvert()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile(__DIR__ . '/../templates/template.xlsx')->toFile('test.csv')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\OneExcelWriter::class, $excel);
    }

    public function testCreateFromFileConvertWithDriver()
    {
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile(__DIR__ . '/../templates/template.xlsx')->toFile('test.csv')->withDriver(Driver::SPOUT)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\SpoutWriter::class, $excel);
    }

    public function testCreateFromFileConvertWithDriverUnsupportedInputOutput()
    {
        $this->expectException(\Exception::class);

        $this->requireLibXL();

        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile(__DIR__ . '/../templates/template.xlsx')->toFile('test.xls')->withDriver(Driver::LIBXL)->make();
    }

    public function testCreateFromFileConvertWithDriverUnsupportedInput()
    {
        $this->expectException(\Exception::class);

        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile(__DIR__ . '/../templates/template.xlsx')->toFile('test.csv')->withDriver(Driver::FPUTCSV)->make();
    }

    public function testCreateFromFileConvertWithDriverUnsupportedOutput()
    {
        $this->expectException(\Exception::class);

        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile(__DIR__ . '/../templates/template.csv')->toFile('test.xlsx')->withDriver(Driver::FPUTCSV)->make();
    }
}
