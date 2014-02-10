<?php

namespace MewesK\TwigExcelBundle\Tests;

use MewesK\TwigExcelBundle\Twig\TwigExcelExtension;
use Twig_Environment;

class TwigTest extends \PHPUnit_Framework_TestCase
{
    private $environment;

    //
    // Helper
    //

    private function getLoaderArray(array $templateNames = array()) {
        $loaderArray = array();
        foreach($templateNames as $templateName) {
            $loaderArray[$templateName] = file_get_contents(__DIR__.'/Resources/views/'.$templateName.'.xls.twig');
        }
        return $loaderArray;
    }

    private function getPhpExcelObject($templateName, $format) {
        // generate source from template
        $source = $this->environment->loadTemplate($templateName)->render(
            array('app' => array('request' => array('requestFormat' => $format)))
        );

        // save source
        file_put_contents(__DIR__.'/Temporary/'.$templateName.'.'.$format, $source);

        // load source
        $reader = null;
        switch($format) {
            case 'xls':
                $reader = new \PHPExcel_Reader_Excel5();
                break;
            case 'xlsx':
                $reader = new \PHPExcel_Reader_Excel2007();
                break;
            default:
                throw new \InvalidArgumentException();
        }
        return $reader->load(__DIR__.'/Temporary/'.$templateName.'.'.$format);
    }

    //
    // PhpUnit
    //

    public function formatProvider() {
        return array(
            array('xls'),
            array('xlsx')
        );
    }

    /**
     * @inheritdoc
     */
    public function setUp() {
        $this->environment = new Twig_Environment(new \Twig_Loader_Array($this->getLoaderArray(array(
                'cellIndex',
                'cellProperties',
                'documentMultiSheet',
                'documentSimple',
                'drawingProperties',
                'drawingSimple',
                'headerFooterComplex',
                'headerFooterDrawing',
                'rowIndex'
            ))), array('strict_variables' => true));
        $this->environment->addExtension(new TwigExcelExtension());
        $this->environment->setCache(__DIR__.'/Temporary/');
    }

    /**
     * @inheritdoc
     */
    protected function tearDown()
    {
        //exec('rm -rf '.__DIR__.'/Temporary/');
    }

    //
    // Tests
    //

    /**
     * @dataProvider formatProvider
     */
    public function testDocumentSimple($format)
    {
        try {
            $phpExcel = $this->getPhpExcelObject('documentSimple', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet does not exist');

            $this->assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            $this->assertEquals('Bar', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            $this->assertEquals('Hello', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            $this->assertEquals('World', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    //
    // Tests depending on testDocumentSimple
    //

    /**
     * @depends testDocumentSimple
     * @dataProvider formatProvider
     */
    public function testCellIndex($format)
    {
        try {
            $phpExcel = $this->getPhpExcelObject('cellIndex', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet does not exist');

            $this->assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            $this->assertNotEquals('Bar',$sheet->getCell('C1')->getValue(),  'Unexpected value in C1');
            $this->assertEquals('Lorem', $sheet->getCell('C1')->getValue(), 'Unexpected value in C1');
            $this->assertEquals('Ipsum', $sheet->getCell('D1')->getValue(), 'Unexpected value in D1');
            $this->assertEquals('Hello', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            $this->assertEquals('World', $sheet->getCell('E1')->getValue(), 'Unexpected value in E1');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    /**
     * @depends testDocumentSimple
     * @dataProvider formatProvider
     */
    public function testCellProperties($format)
    {
        try {
            $phpExcel = $this->getPhpExcelObject('cellProperties', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet does not exist');

            $breaks = $sheet->getBreaks();
            $this->assertCount(1, $breaks, 'Unexpected break count');
            $this->assertArrayHasKey('A1', $breaks, 'Break does not exist');

            $break = $breaks['A1'];
            $this->assertNotNull($break, 'Break is null');

            $cell = $sheet->getCell('A1');
            $this->assertNotNull($cell, 'Cell does not exist');

            $dataValidation = $cell->getDataValidation();
            $this->assertNotNull($dataValidation, 'DataValidation does not exist');

            if ($format != 'xls') {
                $this->assertEquals(true, $dataValidation->getAllowBlank(), 'Unexpected value in allowBlank');
                $this->assertEquals('Test error', $dataValidation->getError(), 'Unexpected value in error');
                $this->assertEquals('information', $dataValidation->getErrorStyle(), 'Unexpected value in errorStyle');
                $this->assertEquals('Test errorTitle', $dataValidation->getErrorTitle(), 'Unexpected value in errorTitle');
                $this->assertEquals('', $dataValidation->getFormula1(), 'Unexpected value in formula1');
                $this->assertEquals('', $dataValidation->getFormula2(), 'Unexpected value in formula2');
                $this->assertEquals('', $dataValidation->getOperator(), 'Unexpected value in operator');
                $this->assertEquals('Test prompt', $dataValidation->getPrompt(), 'Unexpected value in prompt');
                $this->assertEquals('Test promptTitle', $dataValidation->getPromptTitle(), 'Unexpected value in promptTitle');
                $this->assertEquals(true, $dataValidation->getShowDropDown(), 'Unexpected value in showDropDown');
                $this->assertEquals(true, $dataValidation->getShowErrorMessage(), 'Unexpected value in showErrorMessage');
                $this->assertEquals(true, $dataValidation->getShowInputMessage(), 'Unexpected value in showInputMessage');
                $this->assertEquals('custom', $dataValidation->getType(), 'Unexpected value in type');
            }

            $style = $cell->getStyle();
            $this->assertNotNull($style, 'Style does not exist');

            $font = $style->getFont();
            $this->assertNotNull($font, 'Font does not exist');
            $this->assertEquals(18, $font->getSize(), 'Unexpected value in size');

            $this->assertEquals('http://example.com/', $cell->getHyperlink()->getUrl(), 'Unexpected value in url');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    /**
     * @depends testDocumentSimple
     * @dataProvider formatProvider
     */
    public function testDocumentMultiSheet($format)
    {
        try {
            $phpExcel = $this->getPhpExcelObject('documentMultiSheet', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test 1');
            $this->assertNotNull($sheet, 'Sheet "Test 1" does not exist');
            $this->assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');

            $sheet = $phpExcel->getSheetByName('Test 2');
            $this->assertNotNull($sheet, 'Sheet "Test 2" does not exist');
            $this->assertEquals('Bar', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    /**
     * @depends testDocumentSimple
     * @dataProvider formatProvider
     */
    public function testDrawingSimple($format)
    {
        try {
            $phpExcel = $this->getPhpExcelObject('drawingSimple', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet does not exist');

            $drawings = $sheet->getDrawingCollection();
            $this->assertCount(1, $drawings, 'Unexpected drawing count');
            $this->assertArrayHasKey(0, $drawings, 'Drawing does not exist');

            $drawing = $drawings[0];
            $this->assertNotNull($drawing, 'Drawing is null');
            $this->assertEquals(100, $drawing->getWidth(), 'Unexpected value in width');
            $this->assertEquals(100, $drawing->getHeight(), 'Unexpected value in height');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    /**
     * @depends testDocumentSimple
     * @dataProvider formatProvider
     */
    public function testHeaderFooterComplex($format)
    {
        try {
            $phpExcel = $this->getPhpExcelObject('headerFooterComplex', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet does not exist');

            $headerFooter = $sheet->getHeaderFooter();
            $this->assertNotNull($headerFooter, 'HeaderFooter does not exist');
            
            $this->assertEquals('&LoddHeader left&CoddHeader center&RoddHeader right', $headerFooter->getOddHeader(), 'Unexpected value in oddHeader');
            $this->assertEquals('&LoddFooter left&CoddFooter center&RoddFooter right', $headerFooter->getOddFooter(), 'Unexpected value in oddFooter');

            if ($format != 'xls') {
                $this->assertEquals('&LfirstHeader left&CfirstHeader center&RfirstHeader right', $headerFooter->getFirstHeader(), 'Unexpected value in firstHeader');
                $this->assertEquals('&LevenHeader left&CevenHeader center&RevenHeader right', $headerFooter->getEvenHeader(), 'Unexpected value in evenHeader');
                $this->assertEquals('&LfirstFooter left&CfirstFooter center&RfirstFooter right', $headerFooter->getFirstFooter(), 'Unexpected value in firstFooter');
                $this->assertEquals('&LevenFooter left&CevenFooter center&RevenFooter right', $headerFooter->getEvenFooter(), 'Unexpected value in evenFooter');
            }
        }
        catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    /**
     * @depends testDocumentSimple
     * @dataProvider formatProvider
     */
    public function testRowIndex($format)
    {
        try {
            $phpExcel = $this->getPhpExcelObject('rowIndex', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet does not exist');

            $this->assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            $this->assertNotEquals('Bar',$sheet->getCell('A3')->getValue(),  'Unexpected value in A3');
            $this->assertEquals('Lorem', $sheet->getCell('A3')->getValue(), 'Unexpected value in A3');
            $this->assertEquals('Ipsum', $sheet->getCell('A4')->getValue(), 'Unexpected value in A4');
            $this->assertEquals('Hello', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            $this->assertEquals('World', $sheet->getCell('A5')->getValue(), 'Unexpected value in A5');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    //
    // Tests depending on testDrawingSimple
    //

    /**
     * @depends testDrawingSimple
     * @dataProvider formatProvider
     */
    public function testHeaderFooterDrawing($format)
    {
        // header drawings are not supported by the Excel5 writer
        if ($format == 'xls') {
            return;
        }

        try {
            $phpExcel = $this->getPhpExcelObject('headerFooterDrawing', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet does not exist');

            $headerFooter = $sheet->getHeaderFooter();
            $this->assertNotNull($headerFooter, 'HeaderFooter does not exist');
            $this->assertEquals('&L&G&CHeader', $headerFooter->getFirstHeader(), 'Unexpected value in firstHeader');
            $this->assertEquals('&L&G&CHeader', $headerFooter->getEvenHeader(), 'Unexpected value in evenHeader');
            $this->assertEquals('&L&G&CHeader', $headerFooter->getOddHeader(), 'Unexpected value in oddHeader');
            $this->assertEquals('&LFooter&R&G', $headerFooter->getFirstFooter(), 'Unexpected value in firstFooter');
            $this->assertEquals('&LFooter&R&G', $headerFooter->getEvenFooter(), 'Unexpected value in evenFooter');
            $this->assertEquals('&LFooter&R&G', $headerFooter->getOddFooter(), 'Unexpected value in oddFooter');

            $drawings = $headerFooter->getImages();
            $this->assertCount(2, $drawings, 'Sheet has not exactly 2 drawings');
            $this->assertArrayHasKey('LH', $drawings, 'Header drawing does not exist');
            $this->assertArrayHasKey('RF', $drawings, 'Footer drawing does not exist');

            $drawing = $drawings['LH'];
            $this->assertNotNull($drawing, 'Header drawing is null');
            $this->assertEquals(40, $drawing->getWidth(), 'Unexpected value in width');
            $this->assertEquals(40, $drawing->getHeight(), 'Unexpected value in height');

            $drawing = $drawings['RF'];
            $this->assertNotNull($drawing, 'Footer drawing is null');
            $this->assertEquals(20, $drawing->getWidth(), 'Unexpected value in width');
            $this->assertEquals(20, $drawing->getHeight(), 'Unexpected value in height');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    /**
     * @depends testDrawingSimple
     * @dataProvider formatProvider
     */
    public function testDrawingProperties($format)
    {
        try {
            $phpExcel = $this->getPhpExcelObject('drawingProperties', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet does not exist');

            $drawings = $sheet->getDrawingCollection();
            $this->assertCount(1, $drawings, 'Unexpected drawing count');
            $this->assertArrayHasKey(0, $drawings, 'Drawing does not exist');

            $drawing = $drawings[0];
            $this->assertNotNull($drawing, 'Drawing is null');

            $this->assertEquals('B2', $drawing->getCoordinates(), 'Unexpected value in coordinates');
            $this->assertEquals(200, $drawing->getHeight(), 'Unexpected value in height');
            $this->assertEquals(false, $drawing->getResizeProportional(), 'Unexpected value in resizeProportional');
            $this->assertEquals(300, $drawing->getWidth(), 'Unexpected value in width');

            if ($format != 'xls') {
                $this->assertEquals('Test Description', $drawing->getDescription(), 'Unexpected value in description');
                $this->assertEquals('Test Name', $drawing->getName(), 'Unexpected value in name');
                $this->assertEquals(30, $drawing->getOffsetX(), 'Unexpected value in offsetX');
                $this->assertEquals(20, $drawing->getOffsetY(), 'Unexpected value in offsetY');
                $this->assertEquals(45, $drawing->getRotation(), 'Unexpected value in rotation');
            }

            $shadow = $drawing->getShadow();
            $this->assertNotNull($shadow, 'Shadow is null');

            if ($format != 'xls') {
                $this->assertEquals('ctr', $shadow->getAlignment(), 'Unexpected value in alignment');
                $this->assertEquals(100, $shadow->getAlpha(), 'Unexpected value in alpha');
                $this->assertEquals(11, $shadow->getBlurRadius(), 'Unexpected value in blurRadius');
                $this->assertEquals('0000cc', $shadow->getColor()->getRGB(), 'Unexpected value in color');
                $this->assertEquals(30, $shadow->getDirection(), 'Unexpected value in direction');
                $this->assertEquals(4, $shadow->getDistance(), 'Unexpected value in distance');
                $this->assertEquals(true, $shadow->getVisible(), 'Unexpected value in visible');
            }
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }
}
 