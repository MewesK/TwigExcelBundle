<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Tests;

use MewesK\PhpExcelTwigExtensionBundle\Twig\PhpExcelExtension;

class DocumentTest extends \PHPUnit_Framework_TestCase
{
    private $environment;

    public function setUp() {
        $this->environment = new \Twig_Environment(
            new \Twig_Loader_Array([
                'documentSimple' => file_get_contents(__DIR__.'/Resources/views/documentSimple.xls.twig')
            ]),
            ['strict_variables' => true]
        );
        $this->environment->addExtension(new PhpExcelExtension());
        $this->environment->setCache(__DIR__.'/../tmp/');
    }

    public function testDocumentSimple()
    {
        try {
            // generate source from template
            $source = $this->environment->loadTemplate('documentSimple')->render(
                ['app' => ['request' => ['requestFormat' => 'xls']]]
            );

            // save source
            file_put_contents(__DIR__.'/Resources/results/documentSimple.xls',$source);

            // load source
            $reader = new \PHPExcel_Reader_Excel5();
            $phpExcel = $reader->load(__DIR__.'/Resources/results/documentSimple.xls');

            // test
            $sheet = $phpExcel->getSheetByName('Test');
            $this->assertNotNull($sheet);
            $this->assertEquals($sheet->getCell('A1')->getValue(), 'Foo');
            $this->assertEquals($sheet->getCell('B1')->getValue(), 'Bar');
            $this->assertEquals($sheet->getCell('A2')->getValue(), 'Hello');
            $this->assertEquals($sheet->getCell('B2')->getValue(), 'World');
        } catch (\Twig_Error_Runtime $e) {
            var_dump($e);
            $this->fail($e->getMessage());
        }
    }
}
 