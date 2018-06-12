<?php

namespace Tests\VSXLX;

use VSXLSX\Parser;

class ParserTest extends \PHPUnit_Framework_TestCase
{
    public function testParse()
    {
        $inputFile = __DIR__ . '/../__files/workbook_1.xlsx';

        $parser = new Parser($inputFile);
        $parser->parse();
        $parsed = $parser->get_parsed();

        $expectedFirstRow = array(
            'id' => '1',
            'name' => 'foo',
            'number' => '65',
            'date' => '42736',
            'complete' => 'Yes'
        );

        $expectedLastRow = array(
            'id' => '9',
            'name' => '',
            'number' => '44',
            'date' => '42716',
            'complete' => 'No'
        );

        $this->assertCount(9, $parsed);

        foreach ($expectedFirstRow as $key => $value) {
            $this->assertSame($value, $parsed[0][$key]);
        }

        foreach ($expectedLastRow as $key => $value) {
            $this->assertSame($value, $parsed[8][$key]);
        }
    }

    public function testParseWithRowNumbers()
    {
        $inputFile = __DIR__ . '/../__files/workbook_1.xlsx';

        $parser = new Parser($inputFile);
        $parser->row_numbers(true);

        $parser->parse();
        $parsed = $parser->get_parsed();

        $expectedFirstRow = array(
            '__row_number' => 2,
            'id' => '1',
            'name' => 'foo',
            'number' => '65',
            'date' => '42736',
            'complete' => 'Yes'
        );

        $expectedLastRow = array(
            '__row_number' => 10,
            'id' => '9',
            'name' => '',
            'number' => '44',
            'date' => '42716',
            'complete' => 'No'
        );

        $this->assertCount(9, $parsed);

        foreach ($expectedFirstRow as $key => $value) {
            $this->assertSame($value, $parsed[0][$key]);
        }

        foreach ($expectedLastRow as $key => $value) {
            $this->assertSame($value, $parsed[8][$key]);
        }
    }
}
