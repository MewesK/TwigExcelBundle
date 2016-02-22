<?php

namespace MewesK\TwigExcelBundle\Tests\Functional;

use Symfony\Component\HttpFoundation\Response;
use Twig_Error_Runtime;

/**
 * Class DefaultControllerTest
 * @package MewesK\TwigExcelBundle\Tests\Functional
 */
class DefaultControllerTest extends AbstractControllerTest
{
    //
    // PhpUnit
    //

    /**
     * @return array
     */
    public function formatProvider()
    {
        return [['ods'], ['xls'], ['xlsx']];
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
    public function testSimple($format)
    {
        try {
            $document = $this->getDocument(static::$router->generate('test_default', ['templateName' => 'simple', '_format' => $format]), $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals(100270, $sheet->getCell('B22')->getValue(), 'Unexpected value in B22');

            static::assertEquals('=SUM(B2:B21)', $sheet->getCell('B23')->getValue(), 'Unexpected value in B23');
            static::assertTrue($sheet->getCell('B23')->isFormula(), 'Unexpected value in isFormula');
            static::assertEquals(100270, $sheet->getCell('B23')->getCalculatedValue(), 'Unexpected calculated value in B23');
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
    public function testCustomResponse($format)
    {
        try {
            // Generate URI
            $uri = static::$router->generate('test_custom_response', ['templateName' => 'simple', '_format' => $format]);

            // Generate source
            static::$client->request('GET', $uri);

            /**
             * @var $response Response
             */
            $response = static::$client->getResponse();

            static::assertNotNull($response, 'Response does not exist');
            static::assertEquals('attachment; filename="foobar.bin"', $response->headers->get('content-disposition'), 'Unexpected or missing header "Content-Disposition"');
            static::assertEquals('max-age=600, private', $response->headers->get('cache-control'), 'Unexpected or missing header "Cache-Control"');
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
    public function testDocumentTemplatePath1($format)
    {
        try {
            $document = $this->getDocument(static::$router->generate('test_default', ['templateName' => 'documentTemplatePath1', '_format' => $format]), $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheet(0);
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertEquals('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertEquals('Bar', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
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
    public function testDocumentTemplatePath2($format)
    {
        try {
            $document = $this->getDocument(static::$router->generate('test_default', ['templateName' => 'documentTemplatePath2', '_format' => $format]), $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheet(0);
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertEquals('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertEquals('Bar', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }
}
