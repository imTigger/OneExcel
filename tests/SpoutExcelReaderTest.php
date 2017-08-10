<?php

use Imtigger\OneExcel\Driver;
use PHPUnit\Framework\TestCase;

final class SpoutExcelReaderTest extends TestCase {
    public function testReader()
    {
        $excel = \Imtigger\OneExcel\OneExcelReaderFactory::create()->fromFile(__DIR__ . '/01simple.xlsx')->withDriver(Driver::SPOUT)->make();
        $this->assertInstanceOf(\Imtigger\OneExcel\Reader\SpoutReader::class, $excel);

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
                // $this->assertEmpty($row[1]);
                // $this->assertEmpty($row[2]);
                // $this->assertEmpty($row[3]);
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
