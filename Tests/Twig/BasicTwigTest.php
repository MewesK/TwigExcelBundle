<?php

namespace MewesK\TwigExcelBundle\Tests\Twig;

use Twig_Error_Runtime;

/**
 * Class BasicTwigTest
 * @package MewesK\TwigExcelBundle\Tests\Twig
 */
class BasicTwigTest extends AbstractTwigTest
{
    protected static $TEMP_PATH = '/../../tmp/basic/';

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
    public function testBlock($format)
    {
        try {
            $document = $this->getDocument('block', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
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
    public function testBlockOverrideCell($format)
    {
        try {
            $document = $this->getDocument('blockOverrideCell', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertEquals('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertEquals('Bar2', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
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
    public function testBlockOverrideContent($format)
    {
        try {
            $document = $this->getDocument('blockOverrideContent', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertEquals('Foo2', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
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
    public function testBlockOverrideRow($format)
    {
        try {
            $document = $this->getDocument('blockOverrideRow', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello2', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World2', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
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
    public function testBlockOverrideSheet($format)
    {
        try {
            $document = $this->getDocument('blockOverrideSheet', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test2');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello3', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World3', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertNotEquals('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertNotEquals('Bar', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
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
    public function testMacro($format)
    {
        try {
            $document = $this->getDocument('macro', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello1', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World1', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertEquals('Hello2', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertEquals('World2', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');

            $sheet = $document->getSheetByName('Test2');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello3', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World3', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
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
    public function testCellIndex($format)
    {
        try {
            $document = $this->getDocument('cellIndex', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertNotEquals('Bar', $sheet->getCell('C1')->getValue(), 'Unexpected value in C1');
            static::assertEquals('Lorem', $sheet->getCell('C1')->getValue(), 'Unexpected value in C1');
            static::assertEquals('Ipsum', $sheet->getCell('D1')->getValue(), 'Unexpected value in D1');
            static::assertEquals('Hello', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertEquals('World', $sheet->getCell('E1')->getValue(), 'Unexpected value in E1');
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
    public function testCellProperties($format)
    {
        try {
            $document = $this->getDocument('cellProperties', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            $cell = $sheet->getCell('A1');
            static::assertNotNull($cell, 'Cell does not exist');

            $dataValidation = $cell->getDataValidation();
            static::assertNotNull($dataValidation, 'DataValidation does not exist');

            $style = $cell->getStyle();
            static::assertNotNull($style, 'Style does not exist');

            static::assertEquals('s', $sheet->getCell('B1')->getDataType(), 'Unexpected value in dataType');
            static::assertEquals('n', $sheet->getCell('C1')->getDataType(), 'Unexpected value in dataType');
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
    public function testCellFormula($format)
    {
        try {
            $document = $this->getDocument('cellFormula', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('=A1*B1+2', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertTrue($sheet->getCell('A2')->isFormula(), 'Unexpected value in isFormula');
            static::assertEquals(1337, $sheet->getCell('A2')->getCalculatedValue(), 'Unexpected calculated value in A2');

            static::assertEquals('=SUM(A1:B1)', $sheet->getCell('A3')->getValue(), 'Unexpected value in A3');
            static::assertTrue($sheet->getCell('A3')->isFormula(), 'Unexpected value in isFormula');
            static::assertEquals(669.5, $sheet->getCell('A3')->getCalculatedValue(), 'Unexpected calculated value in A3');
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
            static::assertNotNull($document, 'Document does not exist');

            $properties = $document->getProperties();
            static::assertNotNull($properties, 'Properties do not exist');

            // +/- 24h range to allow possible timezone differences (946684800)
            static::assertGreaterThanOrEqual(946598400, $properties->getCreated(), 'Unexpected value in created');
            static::assertLessThanOrEqual(946771200, $properties->getCreated(), 'Unexpected value in created');
            static::assertEquals('Test creator', $properties->getCreator(), 'Unexpected value in creator');

            $defaultStyle = $document->getDefaultStyle();
            static::assertNotNull($defaultStyle, 'DefaultStyle does not exist');

            static::assertEquals('Test description', $properties->getDescription(), 'Unexpected value in description');
            // +/- 24h range to allow possible timezone differences (946684800)
            static::assertGreaterThanOrEqual(946598400, $properties->getModified(), 'Unexpected value in modified');
            static::assertLessThanOrEqual(946771200, $properties->getModified(), 'Unexpected value in modified');

            $security = $document->getSecurity();
            static::assertNotNull($security, 'Security does not exist');

            static::assertEquals('Test subject', $properties->getSubject(), 'Unexpected value in subject');
            static::assertEquals('Test title', $properties->getTitle(), 'Unexpected value in title');
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
    public function testDocumentSimple($format)
    {
        try {
            $document = $this->getDocument('documentSimple', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('Bar', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertEquals('Hello', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertEquals('World', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
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
            $document = $this->getDocument('documentTemplate.' . $format, $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheet(0);
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Hello2', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
            static::assertEquals('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertEquals('Bar2', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
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
    public function testRowIndex($format)
    {
        try {
            $document = $this->getDocument('rowIndex', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertNotEquals('Bar', $sheet->getCell('A3')->getValue(), 'Unexpected value in A3');
            static::assertEquals('Lorem', $sheet->getCell('A3')->getValue(), 'Unexpected value in A3');
            static::assertEquals('Ipsum', $sheet->getCell('A4')->getValue(), 'Unexpected value in A4');
            static::assertEquals('Hello', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
            static::assertEquals('World', $sheet->getCell('A5')->getValue(), 'Unexpected value in A5');
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
    public function testSheetComplex($format)
    {
        try {
            $document = $this->getDocument('sheetComplex', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test 1');
            static::assertNotNull($sheet, 'Sheet "Test 1" does not exist');
            static::assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertEquals('Bar', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');

            $sheet = $document->getSheetByName('Test 2');
            static::assertNotNull($sheet, 'Sheet "Test 2" does not exist');
            static::assertEquals('Hello World', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
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
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            $defaultColumnDimension = $sheet->getDefaultColumnDimension();
            static::assertNotNull($defaultColumnDimension, 'DefaultColumnDimension does not exist');
            static::assertEquals(0, $defaultColumnDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
            static::assertEquals(-1, $defaultColumnDimension->getWidth(), 'Unexpected value in width');
            static::assertEquals(0, $defaultColumnDimension->getXfIndex(), 'Unexpected value in xfIndex');

            $columnDimension = $sheet->getColumnDimension('D');
            static::assertNotNull($columnDimension, 'ColumnDimension does not exist');
            static::assertEquals(0, $columnDimension->getXfIndex(), 'Unexpected value in xfIndex');

            $pageSetup = $sheet->getPageSetup();
            static::assertNotNull($pageSetup, 'PageSetup does not exist');
            static::assertEquals(1, $pageSetup->getFitToHeight(), 'Unexpected value in fitToHeight');
            static::assertFalse($pageSetup->getFitToPage(), 'Unexpected value in fitToPage');
            static::assertEquals(1, $pageSetup->getFitToWidth(), 'Unexpected value in fitToWidth');
            static::assertFalse($pageSetup->getHorizontalCentered(), 'Unexpected value in horizontalCentered');
            static::assertEquals(100, $pageSetup->getScale(), 'Unexpected value in scale');
            static::assertFalse($pageSetup->getVerticalCentered(), 'Unexpected value in verticalCentered');

            $defaultRowDimension = $sheet->getDefaultRowDimension();
            static::assertNotNull($defaultRowDimension, 'DefaultRowDimension does not exist');
            static::assertEquals(0, $defaultRowDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
            static::assertEquals(-1, $defaultRowDimension->getRowHeight(), 'Unexpected value in rowHeight');
            static::assertEquals(0, $defaultRowDimension->getXfIndex(), 'Unexpected value in xfIndex');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }
}
