<?php

namespace MewesK\TwigExcelBundle\Tests\Twig;

use Twig_Error_Runtime;

/**
 * Class CsvTwigTest
 * @package MewesK\TwigExcelBundle\Tests\Twig
 */
class CsvTwigTest extends AbstractTwigTest
{
    protected static $TEMP_PATH = '/../../tmp/csv/';

    //
    // PhpUnit
    //

    /**
     * @return array
     */
    public function formatProvider()
    {
        return [['csv']];
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
    public function testBasic($format)
    {
        try {
            $path = $this->getDocument('documentSimple', $format);

            static::assertTrue(file_exists($path), 'File does not exist');
            static::assertGreaterThan(0, filesize($path), 'File is empty');
            static::assertEquals("\"Foo\",\"Bar\"".PHP_EOL."\"Hello\",\"World\"".PHP_EOL, file_get_contents($path), 'Unexpected content');

        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }

    /**
     * @param string $format
     *
     * @throws \PHPExcel_Exception
     *
     * @dataProvider formatProvider
     */
    public function testDocumentTemplate($format)
    {
        try {
            $path = $this->getDocument('documentTemplate.csv', $format);

            static::assertTrue(file_exists($path), 'File does not exist');
            static::assertGreaterThan(0, filesize($path), 'File is empty');
            static::assertEquals("\"Hello2\",\"World\"".PHP_EOL."\"Foo\",\"Bar2\"".PHP_EOL, file_get_contents($path), 'Unexpected content');

        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }
}
