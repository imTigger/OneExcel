<?php

use Imtigger\OneExcel\Driver;
use PHPUnit\Framework\TestCase;

final class FCsvExcelReaderTest extends TestCase {
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
    public function testReader()
    {
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.xlsx')->withDriver(Driver::LIBXL)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\LibXLReader::class, $excel);

        $i = 1;
        foreach ($excel->row() as $row) {
            if ($i == 1) {
                $this->assertEquals("Hello", $row[0]);
                $this->assertEmpty( $row[1]);
                $this->assertEquals("Hello", $row[2]);
                $this->assertEmpty($row[3]);
            } else if ($i == 2) {
                $this->assertEmpty($row[0]);
                $this->assertEquals("world!", $row[1]);
                $this->assertEmpty( $row[2]);
                $this->assertEquals("world!", $row[3]);
            } else if ($i == 3) {
                $this->assertEmpty($row[0]);
                $this->assertEmpty($row[1]);
                $this->assertEmpty($row[2]);
                $this->assertEmpty($row[3]);
            } else if ($i == 4) {
                $this->assertEquals('Miscellaneous glyphs', $row[0]);
                $this->assertEmpty($row[1]);
                $this->assertEmpty($row[2]);
                $this->assertEmpty($row[3]);
            } else if ($i == 5) {
                $this->assertEquals('éàèùâêîôûëïüÿäöüç', $row[0]);
                $this->assertEmpty($row[1]);
                $this->assertEmpty($row[2]);
                $this->assertEmpty($row[3]);
            }
            $i += 1;
        }
    }
}
