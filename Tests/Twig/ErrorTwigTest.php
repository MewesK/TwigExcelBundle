<?php

namespace MewesK\TwigExcelBundle\Tests\Twig;

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
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testBlockError($format)
    {
        $this->setExpectedException(
            '\Twig_Error_Syntax',
            'Block tags do not work together with Twig tags provided by TwigExcelBundle. Please use \'xlsblock\' instead in "blockError.twig".'
        );
        $this->getDocument('blockError', $format);
    }

    /**
     * @param string $format
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testDocumentError($format)
    {
        $this->setExpectedException(
            '\Twig_Error_Syntax',
            'Node "MewesK\TwigExcelBundle\Twig\Node\XlsDocumentNode" is not allowed inside of Node "MewesK\TwigExcelBundle\Twig\Node\XlsSheetNode"'
        );
        $this->getDocument('documentError', $format);
    }

    /**
     * @param string $format
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testMacroError($format)
    {
        $this->setExpectedException(
            '\Twig_Error_Syntax',
            'Macro tags do not work together with Twig tags provided by TwigExcelBundle. Please use \'xlsmacro\' instead in "macroError.twig".'
        );
        $this->getDocument('macroError', $format);
    }
}
