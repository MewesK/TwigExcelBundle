<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Tests;

use MewesK\PhpExcelTwigExtensionBundle\Twig\PhpExcelExtension;
use Twig_Environment;

class DocumentTest extends \PHPUnit_Framework_TestCase
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
            array('app' => array('request' => array('requestFormat' => 'xls')))
        );

        // save source
        file_put_contents(__DIR__.'/Temporary/'.$templateName.'.xls',$source);

        // load source
        $reader = new \PHPExcel_Reader_Excel5();
        return $reader->load(__DIR__.'/Temporary/'.$templateName.'.xls');
    }

    /**
     * @inheritdoc
     */
    public function setUp() {
        $this->environment = new Twig_Environment(new \Twig_Loader_Array($this->getLoaderArray(array(
                'documentSimple',
                'documentManualIndex'
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

            $this->assertEquals($sheet->getCell('A1')->getValue(), 'Foo', 'A1 does not contain "Foo"');
            $this->assertEquals($sheet->getCell('B1')->getValue(), 'Bar', 'B1 does not contain "Bar"');
            $this->assertEquals($sheet->getCell('A2')->getValue(), 'Hello', 'A2 does not contain "Hello"');
            $this->assertEquals($sheet->getCell('B2')->getValue(), 'World', 'B2 does not contain "World"');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    public function testDocumentManualIndex()
    {
        try {
            $phpExcel = $this->getPhpExcelObject('documentManualIndex');

            // tests
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet, 'Sheet "Test" does not exist');

            $this->assertEquals($sheet->getCell('A1')->getValue(), 'Foo', 'A1 does not contain "Foo"');
            $this->assertEquals($sheet->getCell('C1')->getValue(), 'Bar', 'C1 does not contain "Bar"');
            $this->assertEquals($sheet->getCell('C3')->getValue(), 'Lorem', 'C3 does not contain "Lorem"');
            $this->assertEquals($sheet->getCell('D3')->getValue(), 'Ipsum', 'D3 does not contain "Ipsum"');
            $this->assertEquals($sheet->getCell('B4')->getValue(), 'Hello', 'B4 does not contain "Hello"');
            $this->assertEquals($sheet->getCell('D4')->getValue(), 'World', 'D4 does not contain "World"');
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }
}
 