<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Tests;

use MewesK\PhpExcelTwigExtensionBundle\Twig\PhpExcelExtension;
use Twig_Environment;

class DocumentTest extends \PHPUnit_Framework_TestCase
{
    private $environment;

    public function setUp() {
        $this->environment = new Twig_Environment(
            new \Twig_Loader_Array(array(
                'documentSimple' => file_get_contents(__DIR__.'/Resources/views/documentSimple.xls.twig')
            )),
            array('strict_variables' => true)
        );
        $this->environment->addExtension(new PhpExcelExtension());
        $this->environment->setCache(__DIR__.'/../tmp/');
    }

    public function testDocumentSimple()
    {
        try {
            // generate source from template
            $source = $this->environment->loadTemplate('documentSimple')->render(
                array('app' => array('request' => array('requestFormat' => 'xls')))
            );

            // save source
            file_put_contents(__DIR__.'/Resources/results/documentSimple.xls',$source);

            // load source
            $reader = new \PHPExcel_Reader_Excel5();
            $phpExcel = $reader->load(__DIR__.'/Resources/results/documentSimple.xls');

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
}
 