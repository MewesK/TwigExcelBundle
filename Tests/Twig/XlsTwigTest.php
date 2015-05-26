<?php

namespace MewesK\TwigExcelBundle\Tests\Twig;

use Twig_Error_Runtime;

/**
 * Class XlsTwigTest
 * @package MewesK\TwigExcelBundle\Tests\Twig
 */
class XlsTwigTest extends AbstractTwigTest
{
    protected static $TEMP_PATH = '/../../tmp/xls/';

    //
    // PhpUnit
    //

    /**
     * @return array
     */
    public function formatProvider()
    {
        return [['xls'], ['xlsx']];
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
            $style = $cell->getStyle();

            $breaks = $sheet->getBreaks();
            static::assertCount(1, $breaks, 'Unexpected break count');
            static::assertArrayHasKey('A1', $breaks, 'Break does not exist');

            $break = $breaks['A1'];
            static::assertNotNull($break, 'Break is null');

            $font = $style->getFont();
            static::assertNotNull($font, 'Font does not exist');
            static::assertEquals(18, $font->getSize(), 'Unexpected value in size');

            static::assertEquals('http://example.com/', $cell->getHyperlink()->getUrl(), 'Unexpected value in url');
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
            $defaultStyle = $document->getDefaultStyle();

            static::assertEquals('Test category', $properties->getCategory(), 'Unexpected value in category');

            $font = $defaultStyle->getFont();
            static::assertNotNull($font, 'Font does not exist');
            static::assertEquals(18, $font->getSize(), 'Unexpected value in size');

            static::assertEquals('Test keywords', $properties->getKeywords(), 'Unexpected value in keywords');
            static::assertEquals('Test modifier', $properties->getLastModifiedBy(), 'Unexpected value in lastModifiedBy');
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
            static::assertCount(1, $drawings, 'Unexpected drawing count');
            static::assertArrayHasKey(0, $drawings, 'Drawing does not exist');

            $drawing = $drawings[0];
            static::assertNotNull($drawing, 'Drawing is null');

            static::assertEquals('B2', $drawing->getCoordinates(), 'Unexpected value in coordinates');
            static::assertEquals(200, $drawing->getHeight(), 'Unexpected value in height');
            static::assertFalse($drawing->getResizeProportional(), 'Unexpected value in resizeProportional');
            static::assertEquals(300, $drawing->getWidth(), 'Unexpected value in width');

            $shadow = $drawing->getShadow();
            static::assertNotNull($shadow, 'Shadow is null');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }

    /**
     * @param string $format
     *
     * @dataProvider formatProvider
     */
    public function testDrawingSimple($format)
    {
        try {
            $document = $this->getDocument('drawingSimple', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            $drawings = $sheet->getDrawingCollection();
            static::assertCount(1, $drawings, 'Unexpected drawing count');
            static::assertArrayHasKey(0, $drawings, 'Drawing does not exist');

            $drawing = $drawings[0];
            static::assertNotNull($drawing, 'Drawing is null');
            static::assertEquals(100, $drawing->getWidth(), 'Unexpected value in width');
            static::assertEquals(100, $drawing->getHeight(), 'Unexpected value in height');
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

            static::assertEquals('&LoddHeader left&CoddHeader center&RoddHeader right', $headerFooter->getOddHeader(), 'Unexpected value in oddHeader');
            static::assertEquals('&LoddFooter left&CoddFooter center&RoddFooter right', $headerFooter->getOddFooter(), 'Unexpected value in oddFooter');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }

    /**
     * The following attributes are not supported by the readers and therefore cannot be tested:
     * $columnDimension->getAutoSize() -> false
     * $columnDimension->getCollapsed() -> true
     * $columnDimension->getColumnIndex() -> 1
     * $columnDimension->getVisible() -> false
     * $defaultColumnDimension->getAutoSize() -> true
     * $defaultColumnDimension->getCollapsed() -> false
     * $defaultColumnDimension->getColumnIndex() -> 1
     * $defaultColumnDimension->getVisible() -> true
     * $defaultRowDimension->getCollapsed() -> false
     * $defaultRowDimension->getRowIndex() -> 1
     * $defaultRowDimension->getVisible() -> true
     * $defaultRowDimension->getzeroHeight() -> false
     * $rowDimension->getCollapsed() -> true
     * $rowDimension->getRowIndex() -> 1
     * $rowDimension->getVisible() -> false
     * $rowDimension->getzeroHeight() -> true
     * $sheet->getShowGridlines() -> false
     *
     * @param string $format
     *
     * @throws \PHPExcel_Exception
     *
     * @dataProvider formatProvider
     */
    public function testSheetProperties($format)
    {
        try {
            $document = $this->getDocument('sheetProperties', $format);
            $sheet = $document->getSheetByName('Test');
            $columnDimension = $sheet->getColumnDimension('D');
            $pageSetup = $sheet->getPageSetup();

            static::assertEquals(1, $columnDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
            static::assertEquals(200, $columnDimension->getWidth(), 'Unexpected value in width');

            $pageMargins = $sheet->getPageMargins();
            static::assertNotNull($pageMargins, 'PageMargins does not exist');
            static::assertEquals(1, $pageMargins->getTop(), 'Unexpected value in top');
            static::assertEquals(1, $pageMargins->getBottom(), 'Unexpected value in bottom');
            static::assertEquals(0.75, $pageMargins->getLeft(), 'Unexpected value in left');
            static::assertEquals(0.75, $pageMargins->getRight(), 'Unexpected value in right');
            static::assertEquals(0.5, $pageMargins->getHeader(), 'Unexpected value in header');
            static::assertEquals(0.5, $pageMargins->getFooter(), 'Unexpected value in footer');

            static::assertEquals('landscape', $pageSetup->getOrientation(), 'Unexpected value in orientation');
            static::assertEquals(9, $pageSetup->getPaperSize(), 'Unexpected value in paperSize');
            static::assertEquals('A1:B1', $pageSetup->getPrintArea(), 'Unexpected value in printArea');

            $protection = $sheet->getProtection();
            static::assertTrue($protection->getAutoFilter(), 'Unexpected value in autoFilter');
            static::assertNotNull($protection, 'Protection does not exist');
            static::assertTrue($protection->getDeleteColumns(), 'Unexpected value in deleteColumns');
            static::assertTrue($protection->getDeleteRows(), 'Unexpected value in deleteRows');
            static::assertTrue($protection->getFormatCells(), 'Unexpected value in formatCells');
            static::assertTrue($protection->getFormatColumns(), 'Unexpected value in formatColumns');
            static::assertTrue($protection->getFormatRows(), 'Unexpected value in formatRows');
            static::assertTrue($protection->getInsertColumns(), 'Unexpected value in insertColumns');
            static::assertTrue($protection->getInsertHyperlinks(), 'Unexpected value in insertHyperlinks');
            static::assertTrue($protection->getInsertRows(), 'Unexpected value in insertRows');
            static::assertTrue($protection->getObjects(), 'Unexpected value in objects');
            static::assertEquals(\PHPExcel_Shared_PasswordHasher::hashPassword('testpassword'), $protection->getPassword(), 'Unexpected value in password');
            static::assertTrue($protection->getPivotTables(), 'Unexpected value in pivotTables');
            static::assertTrue($protection->getScenarios(), 'Unexpected value in scenarios');
            static::assertTrue($protection->getSelectLockedCells(), 'Unexpected value in selectLockedCells');
            static::assertTrue($protection->getSelectUnlockedCells(), 'Unexpected value in selectUnlockedCells');
            static::assertTrue($protection->getSheet(), 'Unexpected value in sheet');
            static::assertTrue($protection->getSort(), 'Unexpected value in sort');

            static::assertTrue($sheet->getPrintGridlines(), 'Unexpected value in printGridlines');
            static::assertTrue($sheet->getRightToLeft(), 'Unexpected value in rightToLeft');
            static::assertEquals('c0c0c0', strtolower($sheet->getTabColor()->getRGB()), 'Unexpected value in tabColor');
            static::assertEquals(75, $sheet->getSheetView()->getZoomScale(), 'Unexpected value in zoomScale');

            $rowDimension = $sheet->getRowDimension(2);
            static::assertNotNull($rowDimension, 'RowDimension does not exist');
            static::assertEquals(1, $rowDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
            static::assertEquals(30, $rowDimension->getRowHeight(), 'Unexpected value in rowHeight');
            static::assertEquals(0, $rowDimension->getXfIndex(), 'Unexpected value in xfIndex');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }
}
