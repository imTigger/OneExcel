<?php

use Imtigger\OneExcel\ColumnType;
use Imtigger\OneExcel\Driver;
use PHPUnit\Framework\TestCase;

final class LibXLWriterTest extends TestCase {

    private function getCellValue($filename, $cellName)
    {
        $objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($filename);
        $objExcel = $objReader->load($filename);

        $value = $objExcel->getActiveSheet()->getCell($cellName)->getValue();

        unset($objReader);
        unset($objExcel);

        return $value;
    }

    public function testDependency()
    {
        if (!extension_loaded('excel')) {
            $this->markTestSkipped(
                'The LibXL extension is not available.'
            );
        }
    }

    /**
     * @depends testDependency
     */
    public function testCreate()
    {
        $path = 'tests/test-libxl.xlsx';
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->toFile($path)->withDriver(Driver::LIBXL)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\LibXLWriter::class, $excel);

        $excel->writeCell(1, 0, 'Hello');
        $excel->writeCell(2, 1, 'World');
        $excel->writeCell(3, 2, 3.141592653, ColumnType::NUMERIC);
        $excel->writeRow(4, ['One', 'Excel']);

        $excel->output();

        $this->assertFileExists($path);
        $this->assertGreaterThan(0, filesize($path));

        $this->assertEquals('Hello', $this->getCellValue($path, 'A1'));
        $this->assertEquals('World', $this->getCellValue($path, 'B2'));
        $this->assertEquals(3.141592653, $this->getCellValue($path, 'C3'));
        $this->assertEquals('One', $this->getCellValue($path, 'A4'));
        $this->assertEquals('Excel', $this->getCellValue($path, 'B4'));

        unlink($path);
    }

    /**
     * @depends testDependency
     */
    public function testTemplate()
    {
        $template = __DIR__ . '/../templates/template.xlsx';
        $path = 'tests/test-libxl.xlsx';
        $excel = \Imtigger\OneExcel\OneExcelWriterFactory::create()->fromFile($template)->toFile($path)->withDriver(Driver::LIBXL)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Writer\LibXLWriter::class, $excel);

        $excel->writeCell(2, 0, 'Hello');
        $excel->writeCell(3, 1, 'World');
        $excel->writeCell(4, 2, 3.141592653, ColumnType::NUMERIC);
        $excel->writeRow(5, ['One', 'Excel']);

        $excel->output();

        $this->assertFileExists($path);
        $this->assertGreaterThan(0, filesize($path));

        $this->assertEquals('Title', $this->getCellValue($path, 'A1'));
        $this->assertEquals('Name', $this->getCellValue($path, 'B1'));
        $this->assertEquals('Hello', $this->getCellValue($path, 'A2'));
        $this->assertEquals('World', $this->getCellValue($path, 'B3'));
        $this->assertEquals(3.141592653, $this->getCellValue($path, 'C4'));
        $this->assertEquals('One', $this->getCellValue($path, 'A5'));
        $this->assertEquals('Excel', $this->getCellValue($path, 'B5'));

        unlink($path);
    }
}
