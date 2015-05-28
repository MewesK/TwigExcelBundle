<?php

namespace MewesK\TwigExcelBundle\Tests\Functional;

use Twig_Error_Runtime;

/**
 * Class CustomConfigTest
 * @package MewesK\TwigExcelBundle\Tests\Functional
 */
class CustomConfigTest extends AbstractControllerTest
{
    protected static $CONFIG_FILE = 'config_custom.yml';

    //
    // PhpUnit
    //

    /**
     * @return array
     */
    public function formatProvider()
    {
        return [['xlsx']];
    }

    //
    // Tests
    //

    /**
     * @param string $format
     *
     * @throws \PHPExcel_Exception
     *
     * @dataProvider formatProvider
     */
    public function testCustomConfig($format)
    {
        try {
            $document = $this->getDocument(static::$router->generate('test_default', ['templateName' => 'simple', '_format' => $format]), $format);
            static::assertNotNull($document, 'Document does not exist');

            static::assertFalse(static::$kernel->getContainer()->getParameter('mewes_k_twig_excel.pre_calculate_formulas'), 'Unexpected parameter');
            static::assertStringEndsWith('tmp/phpexcel', static::$kernel->getContainer()->getParameter('mewes_k_twig_excel.disk_caching_directory'), 'Unexpected parameter');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }
}
