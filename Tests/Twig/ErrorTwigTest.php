<?php

namespace MewesK\TwigExcelBundle\Tests\Twig;

use Twig_Error_Runtime;

/**
 * Class ErrorTwigTest
 * @package MewesK\TwigExcelBundle\Tests\Twig
 */
class ErrorTwigTest extends AbstractTwigTest
{
    protected static $TEMP_PATH = '/../../tmp/error/';

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
    public function testBlockError($format)
    {
        $this->setExpectedException(
            '\Twig_Error_Syntax',
            'Block tags do not work together with Twig tags provided by TwigExcelBundle. Please use \'xlsblock\' instead in "blockError.twig".'
        );

        try {
            $this->getDocument('blockError', $format);
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
    public function testDocumentError($format)
    {
        $this->setExpectedException(
            '\Twig_Error_Syntax',
            'Node "MewesK\TwigExcelBundle\Twig\Node\XlsDocumentNode" is not allowed inside of Node "MewesK\TwigExcelBundle\Twig\Node\XlsSheetNode"'
        );

        try {
            $this->getDocument('documentError', $format);
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
    public function testMacroError($format)
    {
        $this->setExpectedException(
            '\Twig_Error_Syntax',
            'Macro tags do not work together with Twig tags provided by TwigExcelBundle. Please use \'xlsmacro\' instead in "macroError.twig".'
        );

        try {
            $this->getDocument('macroError', $format);
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }
}
