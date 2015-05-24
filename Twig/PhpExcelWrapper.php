<?php

namespace MewesK\TwigExcelBundle\Twig;

/**
 * Class PhpExcelWrapper
 *
 * @package MewesK\TwigExcelBundle\Twig
 */
class PhpExcelWrapper
{
    /**
     * @var int
     */
    public static $COLUMN_DEFAULT = 0;
    /**
     * @var int
     */
    public static $ROW_DEFAULT = 1;

    /**
     * @var array
     */
    public $context;
    /**
     * @var \PHPExcel
     */
    public $documentObject;
    /**
     * @var array
     */
    public $documentAttributes;
    /**
     * @var \PHPExcel_Worksheet
     */
    public $sheetObject;
    /**
     * @var array
     */
    public $sheetAttributes;
    /**
     * @var \PHPExcel_Worksheet_HeaderFooter
     */
    public $headerFooterObject;
    /**
     * @var array
     */
    public $headerFooterAttributes;
    /**
     * @var array
     */
    public $alignmentAttributes;
    /**
     * @var \PhpExcel_Cell
     */
    public $cellObject;
    /**
     * @var array
     */
    public $cellAttributes;
    /**
     * @var \PHPExcel_Worksheet_Drawing | \PHPExcel_Worksheet_HeaderFooterDrawing
     */
    public $drawingObject;
    /**
     * @var array
     */
    public $drawingAttributes;
    /**
     * @var int
     */
    public $row;
    /**
     * @var int
     */
    public $column;
    /**
     * @var array
     */
    public $documentMappings;
    /**
     * @var array
     */
    public $sheetMappings;
    /**
     * @var array
     */
    public $footerHeaderMappings;
    /**
     * @var array
     */
    public $cellMappings;
    /**
     * @var array
     */
    public $drawingMappings;

    /**
     * @param array $context
     */
    public function __construct(array $context = [])
    {
        $this->context = $context;

        $this->documentObject = null;
        $this->sheetObject = null;
        $this->headerFooterObject = null;
        $this->cellObject = null;
        $this->drawingObject = null;

        $this->documentAttributes = [];
        $this->sheetAttributes = [];
        $this->headerFooterAttributes = [];
        $this->alignmentAttributes = [];
        $this->cellAttributes = [];
        $this->drawingAttributes = [];
        
        $this->row = null;
        $this->column = null;
        
        $this->documentMappings = [];
        $this->sheetMappings = [];
        $this->footerHeaderMappings = [];
        $this->cellMappings = [];
        $this->drawingMappings = [];

        $this->initDocumentPropertyMappings();
        $this->initSheetPropertyMappings();
        $this->initFooterHeaderPropertyMappings();
        $this->initCellPropertyMappings();
        $this->initDrawingPropertyMappings();
    }

    //
    // Property mappings
    //

    protected function initDocumentPropertyMappings()
    {
        $wrapper = $this; // PHP 5.3 fix

        $this->documentMappings['category'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setCategory($value); };
        $this->documentMappings['company'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setCompany($value); };
        $this->documentMappings['created'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setCreated($value); };
        $this->documentMappings['creator'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setCreator($value); };
        $this->documentMappings['defaultStyle'] = function($value) use ($wrapper) { $wrapper->documentObject->getDefaultStyle()->applyFromArray($value); };
        $this->documentMappings['description'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setDescription($value); };
        $this->documentMappings['format'] = function($value) use ($wrapper) { $wrapper->documentAttributes['format'] = $value; };
        $this->documentMappings['keywords'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setKeywords($value); };
        $this->documentMappings['lastModifiedBy'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setLastModifiedBy($value); };
        $this->documentMappings['manager'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setManager($value); };
        $this->documentMappings['modified'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setModified($value); };
        $this->documentMappings['security']['lockRevision'] = function($value) use ($wrapper) { $wrapper->documentObject->getSecurity()->setLockRevision($value); };
        $this->documentMappings['security']['lockStructure'] = function($value) use ($wrapper) { $wrapper->documentObject->getSecurity()->setLockStructure($value); };
        $this->documentMappings['security']['lockWindows'] = function($value) use ($wrapper) { $wrapper->documentObject->getSecurity()->setLockWindows($value); };
        $this->documentMappings['security']['revisionsPassword'] = function($value) use ($wrapper) { $wrapper->documentObject->getSecurity()->setRevisionsPassword($value); };
        $this->documentMappings['security']['workbookPassword'] = function($value) use ($wrapper) { $wrapper->documentObject->getSecurity()->setWorkbookPassword($value); };
        $this->documentMappings['subject'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setSubject($value); };
        $this->documentMappings['title'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setTitle($value); };
    }

    protected function initSheetPropertyMappings()
    {
        $wrapper = $this; // PHP 5.3 fix

        $this->sheetMappings['columnDimension']['__multi'] = true;
        $this->sheetMappings['columnDimension']['__object'] = function($key) use ($wrapper) { return $key === 'default' ? $wrapper->sheetObject->getDefaultColumnDimension() : $wrapper->sheetObject->getColumnDimension($key); };
        $this->sheetMappings['columnDimension']['autoSize'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['columnDimension']['__object']($key)->setAutoSize($value); };
        $this->sheetMappings['columnDimension']['collapsed'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['columnDimension']['__object']($key)->setCollapsed($value); };
        $this->sheetMappings['columnDimension']['columnIndex'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['columnDimension']['__object']($key)->setColumnIndex($value); };
        $this->sheetMappings['columnDimension']['outlineLevel'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['columnDimension']['__object']($key)->setOutlineLevel($value); };
        $this->sheetMappings['columnDimension']['visible'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['columnDimension']['__object']($key)->setVisible($value); };
        $this->sheetMappings['columnDimension']['width'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['columnDimension']['__object']($key)->setWidth($value); };
        $this->sheetMappings['columnDimension']['xfIndex'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['columnDimension']['__object']($key)->setXfIndex($value); };
        $this->sheetMappings['pageMargins']['top'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageMargins()->setTop($value); };
        $this->sheetMappings['pageMargins']['bottom'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageMargins()->setBottom($value); };
        $this->sheetMappings['pageMargins']['left'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageMargins()->setLeft($value); };
        $this->sheetMappings['pageMargins']['right'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageMargins()->setRight($value); };
        $this->sheetMappings['pageMargins']['header'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageMargins()->setHeader($value); };
        $this->sheetMappings['pageMargins']['footer'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageMargins()->setFooter($value); };
        $this->sheetMappings['pageSetup']['fitToHeight'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageSetup()->setFitToHeight($value); };
        $this->sheetMappings['pageSetup']['fitToPage'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageSetup()->setFitToPage($value); };
        $this->sheetMappings['pageSetup']['fitToWidth'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageSetup()->setFitToWidth($value); };
        $this->sheetMappings['pageSetup']['horizontalCentered'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageSetup()->setHorizontalCentered($value); };
        $this->sheetMappings['pageSetup']['orientation'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageSetup()->setOrientation($value); };
        $this->sheetMappings['pageSetup']['paperSize'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageSetup()->setPaperSize($value); };
        $this->sheetMappings['pageSetup']['printArea'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageSetup()->setPrintArea($value); };
        $this->sheetMappings['pageSetup']['scale'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageSetup()->setScale($value); };
        $this->sheetMappings['pageSetup']['verticalCentered'] = function($value) use ($wrapper) { $wrapper->sheetObject->getPageSetup()->setVerticalCentered($value); };
        $this->sheetMappings['printGridlines'] = function($value) use ($wrapper) { $wrapper->sheetObject->setPrintGridlines($value); };
        $this->sheetMappings['protection']['autoFilter'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setAutoFilter($value); };
        $this->sheetMappings['protection']['deleteColumns'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setDeleteColumns($value); };
        $this->sheetMappings['protection']['deleteRows'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setDeleteRows($value); };
        $this->sheetMappings['protection']['formatCells'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setFormatCells($value); };
        $this->sheetMappings['protection']['formatColumns'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setFormatColumns($value); };
        $this->sheetMappings['protection']['formatRows'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setFormatRows($value); };
        $this->sheetMappings['protection']['insertColumns'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setInsertColumns($value); };
        $this->sheetMappings['protection']['insertHyperlinks'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setInsertHyperlinks($value); };
        $this->sheetMappings['protection']['insertRows'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setInsertRows($value); };
        $this->sheetMappings['protection']['objects'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setObjects($value); };
        $this->sheetMappings['protection']['password'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setPassword($value); };
        $this->sheetMappings['protection']['pivotTables'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setPivotTables($value); };
        $this->sheetMappings['protection']['scenarios'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setScenarios($value); };
        $this->sheetMappings['protection']['selectLockedCells'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setSelectLockedCells($value); };
        $this->sheetMappings['protection']['selectUnlockedCells'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setSelectUnlockedCells($value); };
        $this->sheetMappings['protection']['sheet'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setSheet($value); };
        $this->sheetMappings['protection']['sort'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setSort($value); };
        $this->sheetMappings['rightToLeft'] = function($value) use ($wrapper) { $wrapper->sheetObject->setRightToLeft($value); };
        $this->sheetMappings['rowDimension']['__multi'] = true;
        $this->sheetMappings['rowDimension']['__object'] = function($key) use ($wrapper) { return $key === 'default' ? $wrapper->sheetObject->getDefaultRowDimension() : $wrapper->sheetObject->getRowDimension($key); };
        $this->sheetMappings['rowDimension']['collapsed'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['rowDimension']['__object']($key)->setCollapsed($value); };
        $this->sheetMappings['rowDimension']['outlineLevel'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['rowDimension']['__object']($key)->setOutlineLevel($value); };
        $this->sheetMappings['rowDimension']['rowHeight'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['rowDimension']['__object']($key)->setRowHeight($value); };
        $this->sheetMappings['rowDimension']['rowIndex'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['rowDimension']['__object']($key)->setRowIndex($value); };
        $this->sheetMappings['rowDimension']['visible'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['rowDimension']['__object']($key)->setVisible($value); };
        $this->sheetMappings['rowDimension']['xfIndex'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['rowDimension']['__object']($key)->setXfIndex($value); };
        $this->sheetMappings['rowDimension']['zeroHeight'] = function($key, $value) use ($wrapper) { $wrapper->sheetMappings['rowDimension']['__object']($key)->setZeroHeight($value); };
        $this->sheetMappings['sheetState'] = function($value) use ($wrapper) { $wrapper->sheetObject->setSheetState($value); };
        $this->sheetMappings['showGridlines'] = function($value) use ($wrapper) { $wrapper->sheetObject->setShowGridlines($value); };
        $this->sheetMappings['tabColor'] = function($value) use ($wrapper) { $wrapper->sheetObject->getTabColor()->setRGB($value); };
        $this->sheetMappings['zoomScale'] = function($value) use ($wrapper) { $wrapper->sheetObject->getSheetView()->setZoomScale($value); };
    }

    protected function initFooterHeaderPropertyMappings()
    {
        $wrapper = $this; // PHP 5.3 fix

        $this->footerHeaderMappings['scaleWithDocument'] = function($value) use ($wrapper) { $wrapper->headerFooterObject->setScaleWithDocument($value); };
        $this->footerHeaderMappings['alignWithMargins'] = function($value) use ($wrapper) { $wrapper->headerFooterObject->setAlignWithMargins($value); };
    }

    protected function initCellPropertyMappings()
    {
        $wrapper = $this; // PHP 5.3 fix

        $this->cellMappings['break'] = function($value) use ($wrapper) { $wrapper->sheetObject->setBreak($wrapper->cellObject->getCoordinate(), $value); };
        $this->cellMappings['dataValidation']['allowBlank'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setAllowBlank($value); };
        $this->cellMappings['dataValidation']['error'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setError($value); };
        $this->cellMappings['dataValidation']['errorStyle'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setErrorStyle($value); };
        $this->cellMappings['dataValidation']['errorTitle'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setErrorTitle($value); };
        $this->cellMappings['dataValidation']['formula1'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setFormula1($value); };
        $this->cellMappings['dataValidation']['formula2'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setFormula2($value); };
        $this->cellMappings['dataValidation']['operator'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setOperator($value); };
        $this->cellMappings['dataValidation']['prompt'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setPrompt($value); };
        $this->cellMappings['dataValidation']['promptTitle'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setPromptTitle($value); };
        $this->cellMappings['dataValidation']['showDropDown'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setShowDropDown($value); };
        $this->cellMappings['dataValidation']['showErrorMessage'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setShowErrorMessage($value); };
        $this->cellMappings['dataValidation']['showInputMessage'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setShowInputMessage($value); };
        $this->cellMappings['dataValidation']['type'] = function($value) use ($wrapper) { $wrapper->cellObject->getDataValidation()->setType($value); };
        $this->cellMappings['style'] = function($value) use ($wrapper) { $wrapper->sheetObject->getStyle($wrapper->cellObject->getCoordinate())->applyFromArray($value); };
        $this->cellMappings['url'] = function($value) use ($wrapper) { $wrapper->cellObject->getHyperlink()->setUrl($value); };
    }

    protected function initDrawingPropertyMappings()
    {
        $wrapper = $this; // PHP 5.3 fix

        $this->drawingMappings['coordinates'] = function($value) use ($wrapper) { $wrapper->drawingObject->setCoordinates($value); };
        $this->drawingMappings['description'] = function($value) use ($wrapper) { $wrapper->drawingObject->setDescription($value); };
        $this->drawingMappings['height'] = function($value) use ($wrapper) { $wrapper->drawingObject->setHeight($value); };
        $this->drawingMappings['name'] = function($value) use ($wrapper) { $wrapper->drawingObject->setName($value); };
        $this->drawingMappings['offsetX'] = function($value) use ($wrapper) { $wrapper->drawingObject->setOffsetX($value); };
        $this->drawingMappings['offsetY'] = function($value) use ($wrapper) { $wrapper->drawingObject->setOffsetY($value); };
        $this->drawingMappings['resizeProportional'] = function($value) use ($wrapper) { $wrapper->drawingObject->setResizeProportional($value); };
        $this->drawingMappings['rotation'] = function($value) use ($wrapper) { $wrapper->drawingObject->setRotation($value); };
        $this->drawingMappings['shadow']['alignment'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setAlignment($value); };
        $this->drawingMappings['shadow']['alpha'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setAlpha($value); };
        $this->drawingMappings['shadow']['blurRadius'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setBlurRadius($value); };
        $this->drawingMappings['shadow']['color'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->getColor()->setRGB($value); };
        $this->drawingMappings['shadow']['direction'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setDirection($value); };
        $this->drawingMappings['shadow']['distance'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setDistance($value); };
        $this->drawingMappings['shadow']['visible'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setVisible($value); };
        $this->drawingMappings['width'] = function($value) use ($wrapper) { $wrapper->drawingObject->setWidth($value); };
    }

    //
    // Tags
    //

    /**
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function startDocument(array $properties = null)
    {
        $this->documentObject = new \PHPExcel();
        $this->documentObject->removeSheetByIndex(0);
        $this->documentAttributes['properties'] = $properties ?: [];

        if ($properties !== null) {
            $this->setProperties($properties, $this->documentMappings);
        }
    }

    /**
     * @throws \PHPExcel_Reader_Exception
     */
    public function endDocument()
    {
        // try document property
        if (array_key_exists('format', $this->documentAttributes)) {
            $format = $this->documentAttributes['format'];
        }
        // try symfony request
        else {
            $app = is_array($this->context) && array_key_exists('app', $this->context) ? $this->context['app'] : null;
            $request = $app && is_callable([$app, 'getRequest']) ? $app->getRequest() : null;
            $format = $request && is_callable([$app, 'getRequestFormat']) ? $request->getRequestFormat() : null;
        }
        // set default
        if ($format === null || !is_string($format)) {
            $format = 'xlsx';
        }

        switch(strtolower($format)) {
            case 'csv':
                $writerType = 'CSV';
                break;
            case 'ods':
                $writerType = 'OpenDocument';
                break;
            case 'pdf':
                $writerType = 'PDF';
                break;
            case 'xls':
                $writerType = 'Excel5';
                break;
            case 'xlsx':
                $writerType = 'Excel2007';
                break;
            default:
                throw new \InvalidArgumentException();
        }

        \PHPExcel_IOFactory::createWriter($this->documentObject, $writerType)->save('php://output');

        $this->documentObject = null;
        $this->documentAttributes = [];
    }

    /**
     * @param string $index
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function startSheet($index, array $properties = null)
    {
        if ($index === null || !is_string($index)) {
            throw new \InvalidArgumentException();
        }
        if (!$this->documentObject->sheetNameExists($index)) {
            $this->documentObject->createSheet()->setTitle($index);
        }

        $this->sheetObject = $this->documentObject->setActiveSheetIndexByName($index);
        $this->sheetAttributes['index'] = $index;
        $this->sheetAttributes['properties'] = $properties ?: [];
        
        if ($properties !== null) {
            $this->setProperties($properties, $this->sheetMappings);
        }
    }

    public function endSheet()
    {
        $this->sheetObject = null;
        $this->sheetAttributes = [];
        $this->row = null;
    }

    /**
     * @param null|int $index
     */
    public function startRow($index = null)
    {
        if ($this->sheetObject === null) {
            throw new \LogicException();
        }
        if ($index !== null && !is_int($index)) {
            throw new \InvalidArgumentException();
        }

        $this->row = $index === null ? $this->increaseRow() : $index;
    }

    public function endRow()
    {
        $this->column = null;
    }

    /**
     * @param null|int $index
     * @param null|mixed $value
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function startCell($index = null, $value = null, array $properties = null)
    {
        if ($this->sheetObject === null) {
            throw new \LogicException();
        }
        if ($index !== null && !is_int($index)) {
            throw new \InvalidArgumentException();
        }

        $this->column = $index === null ? $this->increaseColumn() : $index;
        $this->cellObject = $this->sheetObject->getCellByColumnAndRow($this->column, $this->row);

        if ($value !== null) {
            $this->cellObject->setValue($value);
        }
        
        if ($properties !== null) {
            $this->setProperties($properties, $this->cellMappings);
        }

        $this->cellAttributes['value'] = $value;
        $this->cellAttributes['properties'] = $properties ?: [];
    }

    public function endCell()
    {
        $this->cellObject = null;
        $this->cellAttributes = [];
    }

    /**
     * @param string $type
     * @param null|array $properties
     */
    public function startHeaderFooter($type, array $properties = null)
    {
        if ($this->sheetObject === null) {
            throw new \LogicException();
        }
        if (in_array(strtolower($type), ['header', 'oddheader', 'evenheader', 'firstheader', 'footer', 'oddfooter', 'evenfooter', 'firstfooter'], true) === false) {
            throw new \InvalidArgumentException();
        }

        $this->headerFooterObject = $this->sheetObject->getHeaderFooter();
        $this->headerFooterAttributes['value'] = ['left' => null, 'center' => null, 'right' => null]; // will be generated by the alignment tags
        $this->headerFooterAttributes['type'] = $type;
        $this->headerFooterAttributes['properties'] = $properties ?: [];

        if ($properties !== null) {
            $this->setProperties($properties, $this->footerHeaderMappings);
        }
    }

    public function endHeaderFooter()
    {
        $value = implode('', $this->headerFooterAttributes['value']);

        switch(strtolower($this->headerFooterAttributes['type'])) {
            case 'header':
                $this->headerFooterObject->setOddHeader($value);
                $this->headerFooterObject->setEvenHeader($value);
                $this->headerFooterObject->setFirstHeader($value);
                break;
            case 'footer':
                $this->headerFooterObject->setOddFooter($value);
                $this->headerFooterObject->setEvenFooter($value);
                $this->headerFooterObject->setFirstFooter($value);
                break;
            case 'oddheader':
                $this->headerFooterObject->setDifferentOddEven(true);
                $this->headerFooterObject->setOddHeader($value);
                break;
            case 'oddfooter':
                $this->headerFooterObject->setDifferentOddEven(true);
                $this->headerFooterObject->setOddFooter($value);
                break;
            case 'evenheader':
                $this->headerFooterObject->setDifferentOddEven(true);
                $this->headerFooterObject->setEvenHeader($value);
                break;
            case 'evenfooter':
                $this->headerFooterObject->setDifferentOddEven(true);
                $this->headerFooterObject->setEvenFooter($value);
                break;
            case 'firstheader':
                $this->headerFooterObject->setDifferentFirst(true);
                $this->headerFooterObject->setFirstHeader($value);
                break;
            case 'firstfooter':
                $this->headerFooterObject->setDifferentFirst(true);
                $this->headerFooterObject->setFirstFooter($value);
                break;
            default:
                throw new \InvalidArgumentException();
        }

        $this->headerFooterObject = null;
        $this->headerFooterAttributes = [];
    }

    /**
     * @param null|string $type
     * @param null|array $properties
     */
    public function startAlignment($type = null, array $properties = null)
    {
        $this->alignmentAttributes['type'] = $type;
        $this->alignmentAttributes['properties'] = $properties;

        switch(strtolower($this->alignmentAttributes['type'])) {
            case 'left':
                $this->headerFooterAttributes['value']['left'] = '&L';
                break;
            case 'center':
                $this->headerFooterAttributes['value']['center'] = '&C';
                break;
            case 'right':
                $this->headerFooterAttributes['value']['right'] = '&R';
                break;
            default:
                throw new \InvalidArgumentException();
        }
    }

    /**
     * @param null|string $value
     */
    public function endAlignment($value = null)
    {
        switch(strtolower($this->alignmentAttributes['type'])) {
            case 'left':
                if (strpos($this->headerFooterAttributes['value']['left'], '&G') === false) {
                    $this->headerFooterAttributes['value']['left'] .= $value;
                }
                break;
            case 'center':
                if (strpos($this->headerFooterAttributes['value']['center'], '&G') === false) {
                    $this->headerFooterAttributes['value']['center'] .= $value;
                }
                break;
            case 'right':
                if (strpos($this->headerFooterAttributes['value']['right'], '&G') === false) {
                    $this->headerFooterAttributes['value']['right'] .= $value;
                }
                break;
            default:
                throw new \InvalidArgumentException();
        }

        $this->alignmentAttributes = [];
    }

    /**
     * @param string $path
     * @param array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function startDrawing($path, array $properties = null)
    {
        $pathExtension = pathinfo($path, PATHINFO_EXTENSION);
        $tempPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'xlsdrawing' . '_' . md5($path) .
            ($pathExtension ? '.'.$pathExtension : '');

        // make local copy of the asset
        if (!file_exists($tempPath)) {
            $data = file_get_contents($path);
            if ($data === false) {
                throw new \InvalidArgumentException($path . ' does not exist.');
            }
            $temp = fopen($tempPath, 'w+');
            if ($temp === false) {
                throw new \RuntimeException('Cannot open ' . $tempPath);
            }
            fwrite($temp, $data);
            if (fclose($temp) === false) {
                throw new \RuntimeException('Cannot close ' . $tempPath);
            }
            unset($data, $temp);
        }

        if ($this->sheetObject === null) {
            throw new \LogicException();
        }

        if ($this->headerFooterObject) {
            $location = '';
            switch(strtolower($this->alignmentAttributes['type'])) {
                case 'left':
                    $location .= 'L';
                    $this->headerFooterAttributes['value']['left'] .= '&G';
                    break;
                case 'center':
                    $location .= 'C';
                    $this->headerFooterAttributes['value']['center'] .= '&G';
                    break;
                case 'right':
                    $location .= 'R';
                    $this->headerFooterAttributes['value']['right'] .= '&G';
                    break;
                default:
                    throw new \InvalidArgumentException();
            }
            switch(strtolower($this->headerFooterAttributes['type'])) {
                case 'header':
                case 'oddheader':
                case 'evenheader':
                case 'firstheader':
                    $location .= 'H';
                    break;
                case 'footer':
                case 'oddfooter':
                case 'evenfooter':
                case 'firstfooter':
                    $location .= 'F';
                    break;
                default:
                    throw new \InvalidArgumentException();
            }

            $this->drawingObject = new \PHPExcel_Worksheet_HeaderFooterDrawing();
            $this->drawingObject->setPath($tempPath);
            $this->headerFooterObject->addImage($this->drawingObject, $location);
        } else {
            $this->drawingObject = new \PHPExcel_Worksheet_Drawing();
            $this->drawingObject->setWorksheet($this->sheetObject);
            $this->drawingObject->setPath($tempPath);
        }

        if ($properties !== null) {
            $this->setProperties($properties, $this->drawingMappings);
        }
    }

    public function endDrawing()
    {
        $this->drawingObject = null;
        $this->drawingAttributes = [];
    }

    //
    // Helper
    //

    /**
     * @return int|null
     */
    private function increaseRow()
    {
        return $this->row === null ? self::$ROW_DEFAULT : $this->row + 1;
    }

    /**
     * @return int|null
     */
    private function increaseColumn()
    {
        return $this->column === null ? self::$COLUMN_DEFAULT : $this->column + 1;
    }

    /**
     * @param array $properties
     * @param array $mappings
     */
    private function setProperties(array $properties, array $mappings)
    {
        foreach ($properties as $key => $value) {
            if (array_key_exists($key, $mappings)) {
                if (is_array($value) && is_array($mappings) && $key !== 'style' && $key !== 'defaultStyle') {
                    if (array_key_exists('__multi', $mappings[$key]) && $mappings[$key]['__multi'] === true) {
                        foreach ($value as $_key => $_value) {
                            $this->setPropertiesByKey($_key, $_value, $mappings[$key]);
                        }
                    } else {
                        $this->setProperties($value, $mappings[$key]);
                    }
                } else {
                    $mappings[$key]($value);
                }
            }
        }
    }

    /**
     * @param string $key
     * @param array $properties
     * @param array $mappings
     */
    private function setPropertiesByKey($key, array $properties, array $mappings)
    {
        foreach ($properties as $_key => $value) {
            if (array_key_exists($_key, $mappings)) {
                if (is_array($value)) {
                    $this->setPropertiesByKey($key, $value, $mappings[$_key]);
                } else {
                    $mappings[$_key]($key, $value);
                }
            }
        }
    }
}
