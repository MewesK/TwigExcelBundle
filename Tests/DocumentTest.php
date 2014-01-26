<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Tests;

use MewesK\PhpExcelTwigExtensionBundle\Twig\PhpExcelExtension;

class DocumentTest extends \PHPUnit_Framework_TestCase
{
    public function testDocumentSimple()
    {
        $environment = new \Twig_Environment(new \Twig_Loader_Array(
            ['documentSimple' => file_get_contents(__DIR__.'/Resources/views/documentSimple.xls.twig')]
        ), array('strict_variables' => true));

        $environment->addExtension(new PhpExcelExtension());

        $template = $environment->loadTemplate('documentSimple');

        try {
            $source = $template->render([]);
            var_dump($source);
            $this->assertNotEmpty($source);
        } catch (\Twig_Error_Runtime $e) {
            $this->fail($e->getMessage());
        }
    }
}
 