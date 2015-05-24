<?php

namespace MewesK\TwigExcelBundle\Tests;

use InvalidArgumentException;
use MewesK\TwigExcelBundle\Twig\TwigExcelExtension;
use PHPExcel_Reader_Excel2007;
use PHPExcel_Reader_Excel5;
use PHPUnit_Framework_TestCase;
use Twig_Environment;
use Twig_Error_Runtime;
use Twig_Loader_Filesystem;

/**
 * Class TwigTest
 *
 * @package MewesK\TwigExcelBundle\Tests
 */
class TwigTest extends PHPUnit_Framework_TestCase
{
    /**
     * @var Twig_Environment
     */
    protected static $environment;

    //
    // Helper
    //

    /**
     * @param string $templateName
     * @param string $format
     *
     * @return \PHPExcel
     */
    private function getDocument($templateName, $format) {
        // generate source from template
        $source = self::$environment
			->loadTemplate($templateName.'.twig')
			->render(['app' => new MockGlobalVariables($format)]);

        // save source
        file_put_contents(__DIR__.'/Temporary/'.$templateName.'.'.$format, $source);

        // load source
        $reader = null;
        switch($format) {
            case 'xls':
                $reader = new PHPExcel_Reader_Excel5();
                break;
            case 'xlsx':
                $reader = new PHPExcel_Reader_Excel2007();
                break;
            default:
                throw new InvalidArgumentException();
        }
        return $reader->load(__DIR__.'/Temporary/'.$templateName.'.'.$format);
    }

    //
    // PhpUnit
    //

    /**
     * @return array
     */
    public function formatProvider() {
        return [
            ['xls'],
            ['xlsx']
        ];
    }

    /**
     * {@inheritdoc}
     */
    public static function setUpBeforeClass() {
        self::$environment = new Twig_Environment(
            new Twig_Loader_Filesystem(__DIR__.'/Resources/views/'),
            ['strict_variables' => true]
        );
        self::$environment->addExtension(new TwigExcelExtension());
        self::$environment->setCache(__DIR__.'/Temporary/');
    }

    /**
     * {@inheritdoc}
     */
    public static function tearDownAfterClass()
    {
        exec('rm -rf '.__DIR__.'/Temporary/');
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
    public function testCellIndex($format)
    {
        try {
            $document = $this->getDocument('cellIndex', $format);
            static::assertNotNull($document, 'Document does not exist');

            $sheet = $document->getSheetByName('Test');
            static::assertNotNull($sheet, 'Sheet does not exist');

            static::assertEquals('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
            static::assertNotEquals('Bar',$sheet->getCell('C1')->getValue(),  'Unexpected value in C1');
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

            $breaks = $sheet->getBreaks();
            static::assertCount(1, $breaks, 'Unexpected break count');
            static::assertArrayHasKey('A1', $breaks, 'Break does not exist');

            $break = $breaks['A1'];
            static::assertNotNull($break, 'Break is null');

            $cell = $sheet->getCell('A1');
            static::assertNotNull($cell, 'Cell does not exist');

            $dataValidation = $cell->getDataValidation();
            static::assertNotNull($dataValidation, 'DataValidation does not exist');

            if ($format !== 'xls') {
                static::assertEquals(true, $dataValidation->getAllowBlank(), 'Unexpected value in allowBlank');
                static::assertEquals('Test error', $dataValidation->getError(), 'Unexpected value in error');
                static::assertEquals('information', $dataValidation->getErrorStyle(), 'Unexpected value in errorStyle');
                static::assertEquals('Test errorTitle', $dataValidation->getErrorTitle(), 'Unexpected value in errorTitle');
                static::assertEquals('', $dataValidation->getFormula1(), 'Unexpected value in formula1');
                static::assertEquals('', $dataValidation->getFormula2(), 'Unexpected value in formula2');
                static::assertEquals('', $dataValidation->getOperator(), 'Unexpected value in operator');
                static::assertEquals('Test prompt', $dataValidation->getPrompt(), 'Unexpected value in prompt');
                static::assertEquals('Test promptTitle', $dataValidation->getPromptTitle(), 'Unexpected value in promptTitle');
                static::assertEquals(true, $dataValidation->getShowDropDown(), 'Unexpected value in showDropDown');
                static::assertEquals(true, $dataValidation->getShowErrorMessage(), 'Unexpected value in showErrorMessage');
                static::assertEquals(true, $dataValidation->getShowInputMessage(), 'Unexpected value in showInputMessage');
                static::assertEquals('custom', $dataValidation->getType(), 'Unexpected value in type');
            }

            $style = $cell->getStyle();
            static::assertNotNull($style, 'Style does not exist');

            $font = $style->getFont();
            static::assertNotNull($font, 'Font does not exist');
            static::assertEquals(18, $font->getSize(), 'Unexpected value in size');

            static::assertEquals('http://example.com/', $cell->getHyperlink()->getUrl(), 'Unexpected value in url');
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
    public function testDocumentProperties($format)
    {
        try {
            $document = $this->getDocument('documentProperties', $format);
            static::assertNotNull($document, 'Document does not exist');

            $properties = $document->getProperties();
            static::assertNotNull($properties, 'Properties do not exist');

            static::assertEquals('Test category', $properties->getCategory(), 'Unexpected value in category');
            // +/- 24h range to allow possible timezone differences (946684800)
            static::assertGreaterThanOrEqual(946598400, $properties->getCreated(), 'Unexpected value in created');
            static::assertLessThanOrEqual(946771200, $properties->getCreated(), 'Unexpected value in created');
            static::assertEquals('Test creator', $properties->getCreator(), 'Unexpected value in creator');

            $defaultStyle = $document->getDefaultStyle();
            static::assertNotNull($defaultStyle, 'DefaultStyle does not exist');

            $font = $defaultStyle->getFont();
            static::assertNotNull($font, 'Font does not exist');
            static::assertEquals(18, $font->getSize(), 'Unexpected value in size');

            static::assertEquals('Test description', $properties->getDescription(), 'Unexpected value in description');
            static::assertEquals('Test keywords', $properties->getKeywords(), 'Unexpected value in keywords');
            // +/- 24h range to allow possible timezone differences (946684800)
            static::assertGreaterThanOrEqual(946598400, $properties->getModified(), 'Unexpected value in modified');
            static::assertLessThanOrEqual(946771200, $properties->getModified(), 'Unexpected value in modified');
            static::assertEquals('Test modifier', $properties->getLastModifiedBy(), 'Unexpected value in lastModifiedBy');

            $security = $document->getSecurity();
            static::assertNotNull($security, 'Security does not exist');

            // Not supported by the readers - cannot be tested
            //static::assertEquals(true, $security->getLockRevision(), 'Unexpected value in lockRevision');
            //static::assertEquals(true, $security->getLockStructure(), 'Unexpected value in lockStructure');
            //static::assertEquals(true, $security->getLockWindows(), 'Unexpected value in lockWindows');
            //static::assertEquals('test', $security->getRevisionsPassword(), 'Unexpected value in revisionsPassword');
            //static::assertEquals('test', $security->getWorkbookPassword(), 'Unexpected value in workbookPassword');

            static::assertEquals('Test subject', $properties->getSubject(), 'Unexpected value in subject');
            static::assertEquals('Test title', $properties->getTitle(), 'Unexpected value in title');

            if ($format !== 'xls') {
                static::assertEquals('Test company', $properties->getCompany(), 'Unexpected value in company');
                static::assertEquals('Test manager', $properties->getManager(), 'Unexpected value in manager');
            }
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
            static::assertEquals(false, $drawing->getResizeProportional(), 'Unexpected value in resizeProportional');
            static::assertEquals(300, $drawing->getWidth(), 'Unexpected value in width');

            if ($format !== 'xls') {
                static::assertEquals('Test Description', $drawing->getDescription(), 'Unexpected value in description');
                static::assertEquals('Test Name', $drawing->getName(), 'Unexpected value in name');
                static::assertEquals(30, $drawing->getOffsetX(), 'Unexpected value in offsetX');
                static::assertEquals(20, $drawing->getOffsetY(), 'Unexpected value in offsetY');
                static::assertEquals(45, $drawing->getRotation(), 'Unexpected value in rotation');
            }

            $shadow = $drawing->getShadow();
            static::assertNotNull($shadow, 'Shadow is null');

            if ($format !== 'xls') {
                static::assertEquals('ctr', $shadow->getAlignment(), 'Unexpected value in alignment');
                static::assertEquals(100, $shadow->getAlpha(), 'Unexpected value in alpha');
                static::assertEquals(11, $shadow->getBlurRadius(), 'Unexpected value in blurRadius');
                static::assertEquals('0000cc', $shadow->getColor()->getRGB(), 'Unexpected value in color');
                static::assertEquals(30, $shadow->getDirection(), 'Unexpected value in direction');
                static::assertEquals(4, $shadow->getDistance(), 'Unexpected value in distance');
                static::assertEquals(true, $shadow->getVisible(), 'Unexpected value in visible');
            }
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

            if ($format !== 'xls') {
                static::assertEquals('&LfirstHeader left&CfirstHeader center&RfirstHeader right', $headerFooter->getFirstHeader(), 'Unexpected value in firstHeader');
                static::assertEquals('&LevenHeader left&CevenHeader center&RevenHeader right', $headerFooter->getEvenHeader(), 'Unexpected value in evenHeader');
                static::assertEquals('&LfirstFooter left&CfirstFooter center&RfirstFooter right', $headerFooter->getFirstFooter(), 'Unexpected value in firstFooter');
                static::assertEquals('&LevenFooter left&CevenFooter center&RevenFooter right', $headerFooter->getEvenFooter(), 'Unexpected value in evenFooter');
            }
        }
        catch (Twig_Error_Runtime $e) {
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
        // header/footer drawings are not supported by the Excel5 writer
        if ($format === 'xls') {
            return;
        }

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
        // header/footer properties are not supported by the Excel5 writer
        if ($format === 'xls') {
            return;
        }
        
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

            static::assertEquals(false, $headerFooter->getAlignWithMargins(), 'Unexpected value in alignWithMargins');
            static::assertEquals(false, $headerFooter->getScaleWithDocument(), 'Unexpected value in scaleWithDocument');
        }
        catch (Twig_Error_Runtime $e) {
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
            static::assertNotEquals('Bar',$sheet->getCell('A3')->getValue(),  'Unexpected value in A3');
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
            //static::assertEquals(true, $defaultColumnDimension->getAutoSize(), 'Unexpected value in autoSize');
            //static::assertEquals(false, $defaultColumnDimension->getCollapsed(), 'Unexpected value in collapsed');
            //static::assertEquals(1, $defaultColumnDimension->getColumnIndex(), 'Unexpected value in columnIndex');
            static::assertEquals(0, $defaultColumnDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
            //static::assertEquals(true, $defaultColumnDimension->getVisible(), 'Unexpected value in visible');
            static::assertEquals(-1, $defaultColumnDimension->getWidth(), 'Unexpected value in width');
            static::assertEquals(0, $defaultColumnDimension->getXfIndex(), 'Unexpected value in xfIndex');

            $columnDimension = $sheet->getColumnDimension('D');
            static::assertNotNull($columnDimension, 'ColumnDimension does not exist');
            //static::assertEquals(false, $columnDimension->getAutoSize(), 'Unexpected value in autoSize');
            //static::assertEquals(true, $columnDimension->getCollapsed(), 'Unexpected value in collapsed');
            //static::assertEquals(1, $columnDimension->getColumnIndex(), 'Unexpected value in columnIndex');
            static::assertEquals(1, $columnDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
            //static::assertEquals(false, $columnDimension->getVisible(), 'Unexpected value in visible');
            static::assertEquals(200, $columnDimension->getWidth(), 'Unexpected value in width');
            static::assertEquals(0, $columnDimension->getXfIndex(), 'Unexpected value in xfIndex');

            $pageMargins = $sheet->getPageMargins();
            static::assertNotNull($pageMargins, 'PageMargins does not exist');
            static::assertEquals(1, $pageMargins->getTop(), 'Unexpected value in top');
            static::assertEquals(1, $pageMargins->getBottom(), 'Unexpected value in bottom');
            static::assertEquals(0.75, $pageMargins->getLeft(), 'Unexpected value in left');
            static::assertEquals(0.75, $pageMargins->getRight(), 'Unexpected value in right');
            static::assertEquals(0.5, $pageMargins->getHeader(), 'Unexpected value in header');
            static::assertEquals(0.5, $pageMargins->getFooter(), 'Unexpected value in footer');

            $pageSetup = $sheet->getPageSetup();
            static::assertNotNull($pageSetup, 'PageSetup does not exist');
            static::assertEquals(1, $pageSetup->getFitToHeight(), 'Unexpected value in fitToHeight');
            static::assertEquals(false, $pageSetup->getFitToPage(), 'Unexpected value in fitToPage');
            static::assertEquals(1, $pageSetup->getFitToWidth(), 'Unexpected value in fitToWidth');
            static::assertEquals(false, $pageSetup->getHorizontalCentered(), 'Unexpected value in horizontalCentered');
            static::assertEquals('landscape', $pageSetup->getOrientation(), 'Unexpected value in orientation');
            static::assertEquals(9, $pageSetup->getPaperSize(), 'Unexpected value in paperSize');
            static::assertEquals('A1:B1', $pageSetup->getPrintArea(), 'Unexpected value in printArea');
            static::assertEquals(100, $pageSetup->getScale(), 'Unexpected value in scale');
            static::assertEquals(false, $pageSetup->getVerticalCentered(), 'Unexpected value in verticalCentered');

            $protection = $sheet->getProtection();
            static::assertNotNull($protection, 'Protection does not exist');
            static::assertEquals(true, $protection->getAutoFilter(), 'Unexpected value in autoFilter');
            static::assertEquals(true, $protection->getDeleteColumns(), 'Unexpected value in deleteColumns');
            static::assertEquals(true, $protection->getDeleteRows(), 'Unexpected value in deleteRows');
            static::assertEquals(true, $protection->getFormatCells(), 'Unexpected value in formatCells');
            static::assertEquals(true, $protection->getFormatColumns(), 'Unexpected value in formatColumns');
            static::assertEquals(true, $protection->getFormatRows(), 'Unexpected value in formatRows');
            static::assertEquals(true, $protection->getInsertColumns(), 'Unexpected value in insertColumns');
            static::assertEquals(true, $protection->getInsertHyperlinks(), 'Unexpected value in insertHyperlinks');
            static::assertEquals(true, $protection->getInsertRows(), 'Unexpected value in insertRows');
            static::assertEquals(true, $protection->getObjects(), 'Unexpected value in objects');
            static::assertEquals(\PHPExcel_Shared_PasswordHasher::hashPassword('testpassword'), $protection->getPassword(), 'Unexpected value in password');
            static::assertEquals(true, $protection->getPivotTables(), 'Unexpected value in pivotTables');
            static::assertEquals(true, $protection->getScenarios(), 'Unexpected value in scenarios');
            static::assertEquals(true, $protection->getSelectLockedCells(), 'Unexpected value in selectLockedCells');
            static::assertEquals(true, $protection->getSelectUnlockedCells(), 'Unexpected value in selectUnlockedCells');
            static::assertEquals(true, $protection->getSheet(), 'Unexpected value in sheet');
            static::assertEquals(true, $protection->getSort(), 'Unexpected value in sort');

            static::assertEquals(true, $sheet->getPrintGridlines(), 'Unexpected value in printGridlines');
            static::assertEquals(true, $sheet->getRightToLeft(), 'Unexpected value in rightToLeft');

            $defaultRowDimension = $sheet->getDefaultRowDimension();
            static::assertNotNull($defaultRowDimension, 'DefaultRowDimension does not exist');
            //static::assertEquals(false, $defaultRowDimension->getCollapsed(), 'Unexpected value in collapsed');
            static::assertEquals(0, $defaultRowDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
            static::assertEquals(-1, $defaultRowDimension->getRowHeight(), 'Unexpected value in rowHeight');
            //static::assertEquals(1, $defaultRowDimension->getRowIndex(), 'Unexpected value in rowIndex');
            //static::assertEquals(true, $defaultRowDimension->getVisible(), 'Unexpected value in visible');
            static::assertEquals(0, $defaultRowDimension->getXfIndex(), 'Unexpected value in xfIndex');
            //static::assertEquals(false, $defaultRowDimension->getzeroHeight(), 'Unexpected value in zeroHeight');

            $rowDimension = $sheet->getRowDimension(2);
            static::assertNotNull($rowDimension, 'RowDimension does not exist');
            //static::assertEquals(true, $rowDimension->getCollapsed(), 'Unexpected value in collapsed');
            static::assertEquals(1, $rowDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
            static::assertEquals(30, $rowDimension->getRowHeight(), 'Unexpected value in rowHeight');
            //static::assertEquals(1, $rowDimension->getRowIndex(), 'Unexpected value in rowIndex');
            //static::assertEquals(false, $rowDimension->getVisible(), 'Unexpected value in visible');
            static::assertEquals(0, $rowDimension->getXfIndex(), 'Unexpected value in xfIndex');
            //static::assertEquals(true, $rowDimension->getzeroHeight(), 'Unexpected value in zeroHeight');

            // Not supported by the readers - cannot be tested
            //static::assertEquals(false, $sheet->getShowGridlines(), 'Unexpected value in showGridlines');
            static::assertEquals('c0c0c0', strtolower($sheet->getTabColor()->getRGB()), 'Unexpected value in tabColor');
            static::assertEquals(75, $sheet->getSheetView()->getZoomScale(), 'Unexpected value in zoomScale');
        } catch (Twig_Error_Runtime $e) {
            static::fail($e->getMessage());
        }
    }
}
