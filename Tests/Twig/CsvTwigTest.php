<?php

namespace MewesK\TwigExcelBundle\Tests\Twig;

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
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testBasic($format)
    {
        $path = $this->getDocument('documentSimple', $format);

        static::assertFileExists($path, 'File does not exist');
        static::assertGreaterThan(0, filesize($path), 'File is empty');
        static::assertEquals("\"Foo\",\"Bar\"".PHP_EOL."\"Hello\",\"World\"".PHP_EOL, file_get_contents($path), 'Unexpected content');
    }

    /**
     * @param string $format
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testDocumentTemplate($format)
    {
        $path = $this->getDocument('documentTemplate.csv', $format);

        static::assertFileExists($path, 'File does not exist');
        static::assertGreaterThan(0, filesize($path), 'File is empty');
        static::assertEquals("\"Hello2\",\"World\"".PHP_EOL."\"Foo\",\"Bar2\"".PHP_EOL, file_get_contents($path), 'Unexpected content');
    }
}
