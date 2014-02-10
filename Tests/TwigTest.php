<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Tests;

use MewesK\PhpExcelTwigExtensionBundle\Twig\PhpExcelExtension;
use Twig_Environment;

class TwigTest extends \PHPUnit_Framework_TestCase
{
    private $environment;

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
                'documentSimple',
                'documentIndices',
                'drawingSimple',
                'drawingHeaderFooter',
                'documentMultiSheet',
                'drawingProperties'
            ))), array('strict_variables' => true));
        $this->environment->addExtension(new PhpExcelExtension());
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
            $this->assertNotNull($sheet, 'Sheet "Test" does not exist');

            $this->assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'A1 does not equal "Foo"');
            $this->assertEquals('Bar', $sheet->getCell('B1')->getValue(), 'B1 does not equal "Bar"');
            $this->assertEquals('Hello', $sheet->getCell('A2')->getValue(), 'A2 does not equal "Hello"');
            $this->assertEquals('World', $sheet->getCell('B2')->getValue(), 'B2 does not equal "World"');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    /**
     * @depends testDocumentSimple
     * @dataProvider formatProvider
     */
    public function testDocumentManualIndex($format)
    {
        try {
            $phpExcel = $this->getPhpExcelObject('documentIndices', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet "Test" does not exist');

            $this->assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'A1 does not equal "Foo"');
            $this->assertEquals('Bar',$sheet->getCell('C1')->getValue(),  'C1 does not equal "Bar"');
            $this->assertEquals('Lorem', $sheet->getCell('C3')->getValue(), 'C3 does not equal "Lorem"');
            $this->assertEquals('Ipsum', $sheet->getCell('D3')->getValue(), 'D3 does not equal "Ipsum"');
            $this->assertEquals('Hello', $sheet->getCell('B4')->getValue(), 'B4 does not equal "Hello"');
            $this->assertEquals('World', $sheet->getCell('D4')->getValue(), 'D4 does not equal "World"');
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
            $this->assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'A1 does not equal "Foo"');

            $sheet = $phpExcel->getSheetByName('Test 2');
            $this->assertNotNull($sheet, 'Sheet "Test 2" does not exist');
            $this->assertEquals('Bar', $sheet->getCell('A1')->getValue(), 'A1 does not equal "Bar"');
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
            $this->assertNotNull($sheet, 'Sheet "Test" does not exist');

            $drawings = $sheet->getDrawingCollection();
            $this->assertCount(1, $drawings, 'Sheet has not exactly one drawing');
            $this->assertArrayHasKey(0, $drawings, 'Drawing does not exist');

            $drawing = $drawings[0];
            $this->assertNotNull($drawing, 'Drawing is null');
            $this->assertEquals(100, $drawing->getWidth(), 'Drawing width does not equal 100');
            $this->assertEquals(100, $drawing->getHeight(), 'Drawing height does not equal 100');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    /**
     * @depends testDrawingSimple
     * @dataProvider formatProvider
     */
    public function testDrawingHeaderFooter($format)
    {
        // header drawings are not supported by the Excel5 writer
        if ($format == 'xls') {
            return;
        }

        try {
            $phpExcel = $this->getPhpExcelObject('drawingHeaderFooter', $format);

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet "Test" does not exist');

            $drawings = $sheet->getHeaderFooter()->getImages();
            $this->assertCount(2, $drawings, 'Sheet has not exactly 2 drawings');
            $this->assertArrayHasKey('LH', $drawings, 'Header drawing does not exist');
            $this->assertArrayHasKey('RF', $drawings, 'Footer drawing does not exist');

            $drawing = $drawings['LH'];
            $this->assertNotNull($drawing, 'Header drawing is null');
            $this->assertEquals(40, $drawing->getWidth(), 'Width does not equal 40');
            $this->assertEquals(40, $drawing->getHeight(), 'Height does not equal 40');

            $drawing = $drawings['RF'];
            $this->assertNotNull($drawing, 'Footer drawing is null');
            $this->assertEquals(20, $drawing->getWidth(), 'Footer drawing width does not equal 20');
            $this->assertEquals(20, $drawing->getHeight(), 'Footer drawing height does not equal 20');
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
            $this->assertNotNull($sheet, 'Sheet "Test" does not exist');

            $drawings = $sheet->getDrawingCollection();
            $this->assertCount(1, $drawings, 'Sheet has not exactly one drawing');
            $this->assertArrayHasKey(0, $drawings, 'Drawing does not exist');

            $drawing = $drawings[0];
            $this->assertNotNull($drawing, 'Drawing is null');

            $this->assertEquals('B2', $drawing->getCoordinates(), 'Drawing coordinates does not equal "B2"');
            $this->assertEquals(200, $drawing->getHeight(), 'Drawing height does not equal 200');
            $this->assertEquals(false, $drawing->getResizeProportional(), 'Drawing resizeProportional does not equal false');
            $this->assertEquals(300, $drawing->getWidth(), 'Drawing width does not equal 300');

            if ($format != 'xls') {
                $this->assertEquals('Test Description', $drawing->getDescription(), 'Drawing description does not equal "Test Description"');
                $this->assertEquals('Test Name', $drawing->getName(), 'Drawing name does not equal "Test Name"');
                $this->assertEquals(30, $drawing->getOffsetX(), 'Drawing offsetX does not equal 30');
                $this->assertEquals(20, $drawing->getOffsetY(), 'Drawing offsetY does not equal 20');
                $this->assertEquals(45, $drawing->getRotation(), 'Drawing rotation does not equal 45');
            }

            $shadow = $drawing->getShadow();
            $this->assertNotNull($shadow, 'Shadow is null');

            if ($format != 'xls') {
                $this->assertEquals('ctr', $shadow->getAlignment(), 'Shadow alignment does not equal "ctr"');
                $this->assertEquals(100, $shadow->getAlpha(), 'Shadow alpha does not equal 100');
                $this->assertEquals(11, $shadow->getBlurRadius(), 'Shadow blurRadius does not equal 11');
                $this->assertEquals('0000cc', $shadow->getColor()->getRGB(), 'Shadow color does not equal "0000cc"');
                $this->assertEquals(30, $shadow->getDirection(), 'Shadow direction does not equal 30');
                $this->assertEquals(4, $shadow->getDistance(), 'Shadow distance does not equal 4');
                $this->assertEquals(true, $shadow->getVisible(), 'Shadow visible does not equal true');
            }
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }
}
 