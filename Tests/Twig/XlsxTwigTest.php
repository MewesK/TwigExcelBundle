<?php

namespace MewesK\TwigExcelBundle\Tests\Twig;

use Twig_Error_Runtime;

/**
 * Class XlsxTwigTest
 * @package MewesK\TwigExcelBundle\Tests\Twig
 */
class XlsxTwigTest extends AbstractTwigTest
{
    protected static $TEMP_PATH = '/../../tmp/xlsx/';

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
    public function testCellProperties($format)
    {
        try {
            $document = $this->getDocument('cellProperties', $format);
            $sheet = $document->getSheetByName('Test');
            $cell = $sheet->getCell('A1');
            $dataValidation = $cell->getDataValidation();

            static::assertTrue($dataValidation->getAllowBlank(), 'Unexpected value in allowBlank');
            static::assertEquals('Test error', $dataValidation->getError(), 'Unexpected value in error');
            static::assertEquals('information', $dataValidation->getErrorStyle(), 'Unexpected value in errorStyle');
            static::assertEquals('Test errorTitle', $dataValidation->getErrorTitle(), 'Unexpected value in errorTitle');
            static::assertEquals('', $dataValidation->getFormula1(), 'Unexpected value in formula1');
            static::assertEquals('', $dataValidation->getFormula2(), 'Unexpected value in formula2');
            static::assertEquals('', $dataValidation->getOperator(), 'Unexpected value in operator');
            static::assertEquals('Test prompt', $dataValidation->getPrompt(), 'Unexpected value in prompt');
            static::assertEquals('Test promptTitle', $dataValidation->getPromptTitle(), 'Unexpected value in promptTitle');
            static::assertTrue($dataValidation->getShowDropDown(), 'Unexpected value in showDropDown');
            static::assertTrue($dataValidation->getShowErrorMessage(), 'Unexpected value in showErrorMessage');
            static::assertTrue($dataValidation->getShowInputMessage(), 'Unexpected value in showInputMessage');
            static::assertEquals('custom', $dataValidation->getType(), 'Unexpected value in type');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }

    /**
     * The following attributes are not supported by the readers and therefore cannot be tested:
     * $security->getLockRevision() -> true
     * $security->getLockStructure() -> true
     * $security->getLockWindows() -> true
     * $security->getRevisionsPassword() -> 'test'
     * $security->getWorkbookPassword() -> 'test'
     *
     * @param string $format
     *
     * @throws \PHPExcel_Exception
     *
     * @dataProvider formatProvider
     */
    public function testDocumentProperties($format)
    {
        try {
            $document = $this->getDocument('documentProperties', $format);
            $properties = $document->getProperties();

            static::assertEquals('Test company', $properties->getCompany(), 'Unexpected value in company');
            static::assertEquals('Test manager', $properties->getManager(), 'Unexpected value in manager');
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
            $document = $this->getDocument('documentTemplateAdvanced', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheet(0);
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello2', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertEquals('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertEquals('Bar2', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');

            static::assertTrue($sheet->getCell('A1')->getStyle()->getFont()->getBold(), 'Unexpected value in bold');
            static::assertTrue($sheet->getCell('B1')->getStyle()->getFont()->getItalic(), 'Unexpected value in italic');
            static::assertEquals('single', $sheet->getCell('A2')->getStyle()->getFont()->getUnderline(), 'Unexpected value in underline');
            static::assertEquals('FFFF3333', $sheet->getCell('B2')->getStyle()->getFont()->getColor()->getARGB(), 'Unexpected value in color');

            $headerFooter = $sheet->getHeaderFooter();
            static::assertNotNull($headerFooter, 'HeaderFooter does not exist');
            static::assertContains('Left area header',
                $headerFooter->getOddHeader(),
                'Unexpected value in oddHeader');
            static::assertContains('12Center area header',
                $headerFooter->getOddHeader(),
                'Unexpected value in oddHeader');
            static::assertContains('12Right area header',
                $headerFooter->getOddHeader(),
                'Unexpected value in oddHeader');
            static::assertContains('Left area footer',
                $headerFooter->getOddFooter(),
                'Unexpected value in oddFooter');
            static::assertContains('12Center area footer',
                $headerFooter->getOddFooter(),
                'Unexpected value in oddFooter');
            static::assertContains('12Right area footer',
                $headerFooter->getOddFooter(),
                'Unexpected value in oddFooter');

            $drawings = $sheet->getDrawingCollection();
            static::assertCount(1, $drawings, 'Not enough drawings exist');

            $drawing = $drawings[0];
            static::assertEquals(196, $drawing->getWidth(), 'Unexpected value in width');
            static::assertEquals(187, $drawing->getHeight(), 'Unexpected value in height');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }

    /**
     * @param string $format
     *
     * @dataProvider formatProvider
     */
    public function testDrawingProperties($format)
    {
        try {
            $document = $this->getDocument('drawingProperties', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            $drawings = $sheet->getDrawingCollection();
            static::assertCount(1, $drawings, 'Not enough drawings exist');

            $drawing = $drawings[0];
            static::assertEquals('Test Description', $drawing->getDescription(), 'Unexpected value in description');
            static::assertEquals('Test Name', $drawing->getName(), 'Unexpected value in name');
            static::assertEquals(30, $drawing->getOffsetX(), 'Unexpected value in offsetX');
            static::assertEquals(20, $drawing->getOffsetY(), 'Unexpected value in offsetY');
            static::assertEquals(45, $drawing->getRotation(), 'Unexpected value in rotation');

            $shadow = $drawing->getShadow();
            static::assertEquals('ctr', $shadow->getAlignment(), 'Unexpected value in alignment');
            static::assertEquals(100, $shadow->getAlpha(), 'Unexpected value in alpha');
            static::assertEquals(11, $shadow->getBlurRadius(), 'Unexpected value in blurRadius');
            static::assertEquals('0000cc', $shadow->getColor()->getRGB(), 'Unexpected value in color');
            static::assertEquals(30, $shadow->getDirection(), 'Unexpected value in direction');
            static::assertEquals(4, $shadow->getDistance(), 'Unexpected value in distance');
            static::assertTrue($shadow->getVisible(), 'Unexpected value in visible');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }

    /**
     * @param string $format
     *
     * @dataProvider formatProvider
     */
    public function testHeaderFooterComplex($format)
    {
        try {
            $document = $this->getDocument('headerFooterComplex', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            $headerFooter = $sheet->getHeaderFooter();
            static::assertNotNull($headerFooter, 'HeaderFooter does not exist');
            static::assertEquals('&LfirstHeader left&CfirstHeader center&RfirstHeader right',
                $headerFooter->getFirstHeader(),
                'Unexpected value in firstHeader');
            static::assertEquals('&LevenHeader left&CevenHeader center&RevenHeader right',
                $headerFooter->getEvenHeader(),
                'Unexpected value in evenHeader');
            static::assertEquals('&LfirstFooter left&CfirstFooter center&RfirstFooter right',
                $headerFooter->getFirstFooter(),
                'Unexpected value in firstFooter');
            static::assertEquals('&LevenFooter left&CevenFooter center&RevenFooter right',
                $headerFooter->getEvenFooter(),
                'Unexpected value in evenFooter');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }

    /**
     * @param string $format
     *
     * @dataProvider formatProvider
     */
    public function testHeaderFooterDrawing($format)
    {
        try {
            $document = $this->getDocument('headerFooterDrawing', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            $headerFooter = $sheet->getHeaderFooter();
            static::assertNotNull($headerFooter, 'HeaderFooter does not exist');
            static::assertEquals('&L&G&CHeader', $headerFooter->getFirstHeader(), 'Unexpected value in firstHeader');
            static::assertEquals('&L&G&CHeader', $headerFooter->getEvenHeader(), 'Unexpected value in evenHeader');
            static::assertEquals('&L&G&CHeader', $headerFooter->getOddHeader(), 'Unexpected value in oddHeader');
            static::assertEquals('&LFooter&R&G', $headerFooter->getFirstFooter(), 'Unexpected value in firstFooter');
            static::assertEquals('&LFooter&R&G', $headerFooter->getEvenFooter(), 'Unexpected value in evenFooter');
            static::assertEquals('&LFooter&R&G', $headerFooter->getOddFooter(), 'Unexpected value in oddFooter');

            $drawings = $headerFooter->getImages();
            static::assertCount(2, $drawings, 'Sheet has not exactly 2 drawings');
            static::assertArrayHasKey('LH', $drawings, 'Header drawing does not exist');
            static::assertArrayHasKey('RF', $drawings, 'Footer drawing does not exist');

            $drawing = $drawings['LH'];
            static::assertNotNull($drawing, 'Header drawing is null');
            static::assertEquals(40, $drawing->getWidth(), 'Unexpected value in width');
            static::assertEquals(40, $drawing->getHeight(), 'Unexpected value in height');

            $drawing = $drawings['RF'];
            static::assertNotNull($drawing, 'Footer drawing is null');
            static::assertEquals(20, $drawing->getWidth(), 'Unexpected value in width');
            static::assertEquals(20, $drawing->getHeight(), 'Unexpected value in height');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }

    /**
     * @param string $format
     *
     * @dataProvider formatProvider
     */
    public function testHeaderFooterProperties($format)
    {
        try {
            $document = $this->getDocument('headerFooterProperties', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            $headerFooter = $sheet->getHeaderFooter();
            static::assertNotNull($headerFooter, 'HeaderFooter does not exist');

            static::assertEquals('&CHeader', $headerFooter->getFirstHeader(), 'Unexpected value in firstHeader');
            static::assertEquals('&CHeader', $headerFooter->getEvenHeader(), 'Unexpected value in evenHeader');
            static::assertEquals('&CHeader', $headerFooter->getOddHeader(), 'Unexpected value in oddHeader');
            static::assertEquals('&CFooter', $headerFooter->getFirstFooter(), 'Unexpected value in firstFooter');
            static::assertEquals('&CFooter', $headerFooter->getEvenFooter(), 'Unexpected value in evenFooter');
            static::assertEquals('&CFooter', $headerFooter->getOddFooter(), 'Unexpected value in oddFooter');

            static::assertFalse($headerFooter->getAlignWithMargins(), 'Unexpected value in alignWithMargins');
            static::assertFalse($headerFooter->getScaleWithDocument(), 'Unexpected value in scaleWithDocument');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }
}
