<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Tests;

use MewesK\PhpExcelTwigExtensionBundle\Twig\PhpExcelExtension;

class DocumentTest extends \PHPUnit_Framework_TestCase
{
    public function testDocumentSimple()
    {
        $environment = $this->prepareEnvironment([
            'documentSimple' => $this->loadTemplate('documentSimple.xls.twig')
        ]);
        $template = $environment->loadTemplate('documentSimple');

        try {
            $source = $template->render(['app' => ['request' => ['requestFormat' => 'xls']]]);

            //$this->saveResult('documentSimple.xls', $source);

            $this->assertEquals(
                md5($source),
                md5($this->loadReference('documentSimple.xls'))
            );
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }

    public function loadTemplate($name) {
        return file_get_contents(__DIR__.'/Resources/views/'.$name);
    }

    public function loadReference($name) {
        return file_get_contents(__DIR__.'/Resources/references/'.$name);
    }

    public function saveResult($name, $source) {
        file_put_contents(__DIR__.'/Resources/references/'.$name, $source);
    }

    public function prepareEnvironment($templates) {
        $environment = new \Twig_Environment(new \Twig_Loader_Array($templates), ['strict_variables' => true]);
        $environment->addExtension(new PhpExcelExtension());

        //$environment->setCache(__DIR__.'/../tmp/');

        return $environment;
    }
}
 