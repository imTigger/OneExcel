<?php

use Imtigger\OneExcel\Driver;
use Imtigger\OneExcel\OneExcelReaderFactory;
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
        $excel = OneExcelReaderFactory::create();
        $this->assertInstanceOf(\Imtigger\OneExcel\OneExcelReaderFactory::class, $excel);
    }

    public function testCreateXLSX()
    {
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.xlsx')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\OneExcelReader::class, $excel);
    }

    public function testCreateXLS()
    {
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.xls')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\OneExcelReader::class, $excel);
    }

    public function testCreateCSV()
    {
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.csv')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\OneExcelReader::class, $excel);
    }

    public function testCreateODS()
    {
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.ods')->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\OneExcelReader::class, $excel);
    }


    public function testCreatePHPExcel()
    {
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.xlsx')->withDriver(Driver::PHPEXCEL)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\PHPExcelReader::class, $excel);
    }

    public function testCreateLibXL()
    {
        $this->requireLibXL();

        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.xlsx')->withDriver(Driver::LIBXL)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\LibXLReader::class, $excel);
    }

    public function testCreateSpout()
    {
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.xlsx')->withDriver(Driver::SPOUT)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\SpoutReader::class, $excel);
    }

    public function testCreateFputcsv()
    {
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.csv')->withDriver(Driver::FPUTCSV)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\FCsvReader::class, $excel);
    }

    public function testUnsupportedDriver()
    {
        $this->expectException(\Exception::class);
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.xlsx')->withDriver(Driver::FPUTCSV)->make();
    }
}
