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

    private function getPhpExcelObject($templateName) {
        // generate source from template
        $source = $this->environment->loadTemplate($templateName)->render(
            array('app' => array('request' => array('requestFormat' => 'xlsx')))
        );

        // save source
        file_put_contents(__DIR__.'/Temporary/'.$templateName.'.xlsx', $source);

        // load source
        $reader = new \PHPExcel_Reader_Excel2007();
        return $reader->load(__DIR__.'/Temporary/'.$templateName.'.xlsx');
    }

    /**
     * @inheritdoc
     */
    public function setUp() {
        $this->environment = new Twig_Environment(new \Twig_Loader_Array($this->getLoaderArray(array(
                'documentSimple',
                'documentIndices',
                'drawingSimple',
                'drawingHeaderFooter'
            ))), array('strict_variables' => true));
        $this->environment->addExtension(new PhpExcelExtension());
        $this->environment->setCache(__DIR__.'/Temporary/');
    }

    /**
     * @inheritdoc
     */
    protected function tearDown()
    {
        exec('rm -rf '.__DIR__.'/Temporary/');
    }

    public function testDocumentSimple()
    {
        try {
            $phpExcel = $this->getPhpExcelObject('documentSimple');

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

    public function testDocumentManualIndex()
    {
        try {
            $phpExcel = $this->getPhpExcelObject('documentIndices');

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

    public function testDrawingSimple()
    {
        try {
            $phpExcel = $this->getPhpExcelObject('drawingSimple');

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet "Test" does not exist');

            $drawings = $sheet->getDrawingCollection();
            $this->assertCount(1, $drawings, 'Sheet has not exactly one drawing');
            $this->assertArrayHasKey(0, $drawings, 'Drawing does not exist');

            $drawing = $drawings[0];
            $this->assertNotNull($drawing, 'Drawing is null');
            $this->assertEquals(100, $drawing->getWidth(), 'Drawing width does not equal 100px');
            $this->assertEquals(100, $drawing->getHeight(), 'Drawing height does not equal 100px');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    public function testDrawingHeaderFooter()
    {
        try {
            $phpExcel = $this->getPhpExcelObject('drawingHeaderFooter');

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet "Test" does not exist');

            $drawings = $sheet->getHeaderFooter()->getImages();
            $this->assertCount(2, $drawings, 'Sheet has not exactly 2 drawings');
            $this->assertArrayHasKey('LH', $drawings, 'Header drawing does not exist');
            $this->assertArrayHasKey('RF', $drawings, 'Footer drawing does not exist');

            $drawing = $drawings['LH'];
            $this->assertNotNull($drawing, 'Header drawing is null');
            $this->assertEquals(40, $drawing->getWidth(), 'Width does not equal 40px');
            $this->assertEquals(40, $drawing->getHeight(), 'Height does not equal 40px');

            $drawing = $drawings['RF'];
            $this->assertNotNull($drawing, 'Footer drawing is null');
            $this->assertEquals(20, $drawing->getWidth(), 'Footer drawing width does not equal 20px');
            $this->assertEquals(20, $drawing->getHeight(), 'Footer drawing height does not equal 20px');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }
}
 