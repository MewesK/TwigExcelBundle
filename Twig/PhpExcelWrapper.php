<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig;

class PhpExcelWrapper {

    /**
     * @var string
     */
    public static $COLUMN_DEFAULT = 0;
    /**
     * @var int
     */
    public static $ROW_DEFAULT = 1;

    /**
     * @var \PHPExcel
     */
    protected $documentObject;
    /**
     * @var \PHPExcel_Worksheet
     */
    public $sheetObject;
    /**
     * @var \PhpExcel_Cell
     */
    protected $cellObject;
    /**
     * @var \PHPExcel_Worksheet_Drawing
     */
    protected $drawingObject;
    /**
     * @var int
     */
    protected $row;
    /**
     * @var string
     */
    protected $column;
    /**
     * @var string
     */
    protected $format;

    /**
     * @var array
     */
    protected $documentMappings;
    /**
     * @var array
     */
    protected $sheetMappings;
    /**
     * @var array
     */
    protected $rowMappings;
    /**
     * @var array
     */
    protected $cellMappings;
    /**
     * @var array
     */
    protected $drawingMappings;
    

    public function __construct() {
        $this->documentObject = new \PHPExcel();
        $this->documentObject->removeSheetByIndex(0);

        $this->sheetObject = null;
        $this->cellObject = null;
        $this->drawingObject = null;
        $this->row = null;
        $this->column = null;
        $this->format = null;

        $this->documentMappings = [];
        $this->sheetMappings = [];
        $this->rowMappings = [];
        $this->cellMappings = [];
        $this->drawingMappings = [];

        $this->initDocumentPropertyMappings();
        $this->initSheetPropertyMappings();
        $this->initRowPropertyMappings();
        $this->initCellPropertyMappings();
        $this->initDrawingPropertyMappings();
    }

    protected function initDocumentPropertyMappings() {
        $this->documentMappings['category'] = function($value) { $this->documentObject->getProperties()->setCategory($value); };
        $this->documentMappings['company'] = function($value) { $this->documentObject->getProperties()->setCompany($value); };
        $this->documentMappings['created'] = function($value) { $this->documentObject->getProperties()->setCreated($value); };
        $this->documentMappings['creator'] = function($value) { $this->documentObject->getProperties()->setCreator($value); };
        $this->documentMappings['defaultStyle'] = function($value) { $this->documentObject->getDefaultStyle()->applyFromArray($value); };
        $this->documentMappings['description'] = function($value) { $this->documentObject->getProperties()->setDescription($value); };
        $this->documentMappings['format'] = function($value) { $this->format = $value; };
        $this->documentMappings['keywords'] = function($value) { $this->documentObject->getProperties()->setKeywords($value); };
        $this->documentMappings['lastModifiedBy'] = function($value) { $this->documentObject->getProperties()->setLastModifiedBy($value); };
        $this->documentMappings['manager'] = function($value) { $this->documentObject->getProperties()->setManager($value); };
        $this->documentMappings['modified'] = function($value) { $this->documentObject->getProperties()->setModified($value); };
        $this->documentMappings['security']['lockRevision'] = function($value) { $this->documentObject->getSecurity()->setLockRevision($value); };
        $this->documentMappings['security']['lockStructure'] = function($value) { $this->documentObject->getSecurity()->setLockStructure($value); };
        $this->documentMappings['security']['lockWindows'] = function($value) { $this->documentObject->getSecurity()->setLockWindows($value); };
        $this->documentMappings['security']['revisionsPassword'] = function($value) { $this->documentObject->getSecurity()->setRevisionsPassword($value); };
        $this->documentMappings['security']['workbookPassword'] = function($value) { $this->documentObject->getSecurity()->setWorkbookPassword($value); };
        $this->documentMappings['subject'] = function($value) { $this->documentObject->getProperties()->setSubject($value); };
        $this->documentMappings['title'] = function($value) { $this->documentObject->getProperties()->setTitle($value); };
    }

    protected function initSheetPropertyMappings() {
        $this->sheetMappings['columnDimension']['__multi'] = true;
        $this->sheetMappings['columnDimension']['__object'] = function($key) { return $key == 'default' ? $this->sheetObject->getDefaultColumnDimension() : $this->sheetObject->getColumnDimension($key); };
        $this->sheetMappings['columnDimension']['autoSize'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setAutoSize($value); };
        $this->sheetMappings['columnDimension']['collapsed'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setCollapsed($value); };
        $this->sheetMappings['columnDimension']['columnIndex'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setColumnIndex($value); };
        $this->sheetMappings['columnDimension']['outlineLevel'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setOutlineLevel($value); };
        $this->sheetMappings['columnDimension']['visible'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setVisible($value); };
        $this->sheetMappings['columnDimension']['width'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setWidth($value); };
        $this->sheetMappings['columnDimension']['xfIndex'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setXfIndex($value); };
        $this->sheetMappings['footer'] = function($value) { $this->sheetObject->getHeaderFooter()->setOddFooter($value); };
        $this->sheetMappings['header'] = function($value) { $this->sheetObject->getHeaderFooter()->setOddHeader($value); };
        $this->sheetMappings['pageMargins']['top'] = function($value) { $this->sheetObject->getPageMargins()->setTop($value); };
        $this->sheetMappings['pageMargins']['bottom'] = function($value) { $this->sheetObject->getPageMargins()->setBottom($value); };
        $this->sheetMappings['pageMargins']['left'] = function($value) { $this->sheetObject->getPageMargins()->setLeft($value); };
        $this->sheetMappings['pageMargins']['right'] = function($value) { $this->sheetObject->getPageMargins()->setRight($value); };
        $this->sheetMappings['pageMargins']['header'] = function($value) { $this->sheetObject->getPageMargins()->setHeader($value); };
        $this->sheetMappings['pageMargins']['footer'] = function($value) { $this->sheetObject->getPageMargins()->setFooter($value); };
        $this->sheetMappings['pageSetup']['fitToHeight'] = function($value) { $this->sheetObject->getPageSetup()->setFitToHeight($value); };
        $this->sheetMappings['pageSetup']['fitToPage'] = function($value) { $this->sheetObject->getPageSetup()->setFitToPage($value); };
        $this->sheetMappings['pageSetup']['fitToWidth'] = function($value) { $this->sheetObject->getPageSetup()->setFitToWidth($value); };
        $this->sheetMappings['pageSetup']['horizontalCentered'] = function($value) { $this->sheetObject->getPageSetup()->setHorizontalCentered($value); };
        $this->sheetMappings['pageSetup']['orientation'] = function($value) { $this->sheetObject->getPageSetup()->setOrientation($value); };
        $this->sheetMappings['pageSetup']['paperSize'] = function($value) { $this->sheetObject->getPageSetup()->setPaperSize($value); };
        $this->sheetMappings['pageSetup']['printArea'] = function($value) { $this->sheetObject->getPageSetup()->setPrintArea($value); };
        $this->sheetMappings['pageSetup']['scale'] = function($value) { $this->sheetObject->getPageSetup()->setScale($value); };
        $this->sheetMappings['pageSetup']['verticalCentered'] = function($value) { $this->sheetObject->getPageSetup()->setVerticalCentered($value); };
        $this->sheetMappings['printGridlines'] = function($value) { $this->sheetObject->setPrintGridlines($value); };
        $this->sheetMappings['protection']['autoFilter'] = function($value) { $this->sheetObject->getProtection()->setAutoFilter($value); };
        $this->sheetMappings['protection']['deleteColumns'] = function($value) { $this->sheetObject->getProtection()->setDeleteColumns($value); };
        $this->sheetMappings['protection']['deleteRows'] = function($value) { $this->sheetObject->getProtection()->setDeleteRows($value); };
        $this->sheetMappings['protection']['formatCells'] = function($value) { $this->sheetObject->getProtection()->setFormatCells($value); };
        $this->sheetMappings['protection']['formatColumns'] = function($value) { $this->sheetObject->getProtection()->setFormatColumns($value); };
        $this->sheetMappings['protection']['formatRows'] = function($value) { $this->sheetObject->getProtection()->setFormatRows($value); };
        $this->sheetMappings['protection']['insertColumns'] = function($value) { $this->sheetObject->getProtection()->setInsertColumns($value); };
        $this->sheetMappings['protection']['insertHyperlinks'] = function($value) { $this->sheetObject->getProtection()->setInsertHyperlinks($value); };
        $this->sheetMappings['protection']['insertRows'] = function($value) { $this->sheetObject->getProtection()->setInsertRows($value); };
        $this->sheetMappings['protection']['objects'] = function($value) { $this->sheetObject->getProtection()->setObjects($value); };
        $this->sheetMappings['protection']['pivotTables'] = function($value) { $this->sheetObject->getProtection()->setPivotTables($value); };
        $this->sheetMappings['protection']['scenarios'] = function($value) { $this->sheetObject->getProtection()->setScenarios($value); };
        $this->sheetMappings['protection']['selectLockedCells'] = function($value) { $this->sheetObject->getProtection()->setSelectLockedCells($value); };
        $this->sheetMappings['protection']['selectUnlockedCells'] = function($value) { $this->sheetObject->getProtection()->setSelectUnlockedCells($value); };
        $this->sheetMappings['protection']['sheet'] = function($value) { $this->sheetObject->getProtection()->setSheet($value); };
        $this->sheetMappings['protection']['sort'] = function($value) { $this->sheetObject->getProtection()->setSort($value); };
        $this->sheetMappings['rightToLeft'] = function($value) { $this->sheetObject->setRightToLeft($value); };
        $this->sheetMappings['rowDimension']['__multi'] = true;
        $this->sheetMappings['rowDimension']['__object'] = function($key) { return $key == 'default' ? $this->sheetObject->getDefaultRowDimension() : $this->sheetObject->getRowDimension($key); };
        $this->sheetMappings['rowDimension']['collapsed'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setCollapsed($value); };
        $this->sheetMappings['rowDimension']['outlineLevel'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setOutlineLevel($value); };
        $this->sheetMappings['rowDimension']['rowHeight'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setRowHeight($value); };
        $this->sheetMappings['rowDimension']['rowIndex'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setRowIndex($value); };
        $this->sheetMappings['rowDimension']['visible'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setVisible($value); };
        $this->sheetMappings['rowDimension']['xfIndex'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setXfIndex($value); };
        $this->sheetMappings['rowDimension']['zeroHeight'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setZeroHeight($value); };
        $this->sheetMappings['sheetState'] = function($value) { $this->sheetObject->setSheetState($value); };
        $this->sheetMappings['showGridlines'] = function($value) { $this->sheetObject->setShowGridlines($value); };
        $this->sheetMappings['tabColor'] = function($value) { $this->sheetObject->getTabColor()->setRGB($value); };
        $this->sheetMappings['zoomScale'] = function($value) { $this->sheetObject->getSheetView()->setZoomScale($value); };
    }

    protected function initRowPropertyMappings() {
        // nothing
    }

    protected function initCellPropertyMappings() {
        $this->cellMappings['break'] = function($value) { $this->sheetObject->setBreak($this->cellObject->getCoordinate(), $value); };
        $this->cellMappings['dataValidation']['allowBlank'] = function($value) { $this->cellObject->getDataValidation()->setAllowBlank($value); };
        $this->cellMappings['dataValidation']['error'] = function($value) { $this->cellObject->getDataValidation()->setError($value); };
        $this->cellMappings['dataValidation']['errorStyle'] = function($value) { $this->cellObject->getDataValidation()->setErrorStyle($value); };
        $this->cellMappings['dataValidation']['errorTitle'] = function($value) { $this->cellObject->getDataValidation()->setErrorTitle($value); };
        $this->cellMappings['dataValidation']['formula1'] = function($value) { $this->cellObject->getDataValidation()->setFormula1($value); };
        $this->cellMappings['dataValidation']['formula2'] = function($value) { $this->cellObject->getDataValidation()->setFormula2($value); };
        $this->cellMappings['dataValidation']['operator'] = function($value) { $this->cellObject->getDataValidation()->setOperator($value); };
        $this->cellMappings['dataValidation']['prompt'] = function($value) { $this->cellObject->getDataValidation()->setPrompt($value); };
        $this->cellMappings['dataValidation']['promptTitle'] = function($value) { $this->cellObject->getDataValidation()->setPromptTitle($value); };
        $this->cellMappings['dataValidation']['showDropDown'] = function($value) { $this->cellObject->getDataValidation()->setShowDropDown($value); };
        $this->cellMappings['dataValidation']['showErrorMessage'] = function($value) { $this->cellObject->getDataValidation()->setShowErrorMessage($value); };
        $this->cellMappings['dataValidation']['showInputMessage'] = function($value) { $this->cellObject->getDataValidation()->setShowInputMessage($value); };
        $this->cellMappings['dataValidation']['type'] = function($value) { $this->cellObject->getDataValidation()->setType($value); };
        $this->cellMappings['style'] = function($value) { $this->sheetObject->getStyle($this->cellObject->getCoordinate())->applyFromArray($value); };
        $this->cellMappings['url'] = function($value) { $this->cellObject->getHyperlink()->setUrl($value); };
    }

    protected function initDrawingPropertyMappings() {
        $this->drawingMappings['coordinates'] = function($value) { $this->drawingObject->getCoordinates($value); };
        $this->drawingMappings['description'] = function($value) { $this->drawingObject->getDescription($value); };
        $this->drawingMappings['height'] = function($value) { $this->drawingObject->getHeight($value); };
        $this->drawingMappings['name'] = function($value) { $this->drawingObject->getName($value); };
        $this->drawingMappings['offsetX'] = function($value) { $this->drawingObject->getOffsetX($value); };
        $this->drawingMappings['offsetY'] = function($value) { $this->drawingObject->getOffsetY($value); };
        $this->drawingMappings['resizeProportional'] = function($value) { $this->drawingObject->getResizeProportional($value); };
        $this->drawingMappings['rotation'] = function($value) { $this->drawingObject->getRotation($value); };
        $this->drawingMappings['shadow']['alignment'] = function($value) { $this->drawingObject->getShadow()->setAlignment($value); };
        $this->drawingMappings['shadow']['alpha'] = function($value) { $this->drawingObject->getShadow()->setAlpha($value); };
        $this->drawingMappings['shadow']['blurRadius'] = function($value) { $this->drawingObject->getShadow()->setBlurRadius($value); };
        $this->drawingMappings['shadow']['color'] = function($value) { $this->drawingObject->getShadow()->getColor()->setRgb($value); };
        $this->drawingMappings['shadow']['direction'] = function($value) { $this->drawingObject->getShadow()->setDirection($value); };
        $this->drawingMappings['shadow']['distance'] = function($value) { $this->drawingObject->getShadow()->setDistance($value); };
        $this->drawingMappings['shadow']['visible'] = function($value) { $this->drawingObject->getShadow()->setVisible($value); };
        $this->drawingMappings['width'] = function($value) { $this->drawingObject->getWidth($value); };
    }

    public function setProperties(array $properties, array $mappings) {
        foreach ($properties as $key => $value) {
            if (array_key_exists($key, $mappings)) {
                if (is_array($value)) {
                    if ($mappings['__multi']) {
                        foreach ($value as $_key => $_value) {
                            $this->setPropertiesByKey($_key, $_value, $mappings[$key]);
                        }
                    } else {
                        $this->setProperties($value, $mappings[$key]);
                    }
                } else {
                    $mappings[$key](trim($value));
                }
            }
        }
    }

    public function setPropertiesByKey($key, array $properties, array $mappings) {
        foreach ($properties as $_key => $value) {
            if (array_key_exists($_key, $mappings)) {
                if (is_array($value)) {
                    $this->setPropertiesByKey($key, $value, $mappings[$_key]);
                } else {
                    $mappings[$_key]($key, trim($value));
                }
            }
        }
    }

    public function tagDocument(array $properties = null) {
        if ($properties != null) {
            $this->setProperties($properties, $this->documentMappings);
        }
    }

    public function tagSheet($index, array $properties = null) {
        if ($index == null || empty($index)) {
            throw new \InvalidArgumentException();
        }

        if (!$this->documentObject->sheetNameExists($index)) {
            $this->documentObject->createSheet()->setTitle($index);
        }

        $this->column = null;
        $this->row = null;

        $this->sheetObject = $this->documentObject->setActiveSheetIndexByName($index);
        $this->cellObject = null;
        
        if ($properties != null) {
            $this->setProperties($properties, $this->sheetMappings);
        }
    }

    public function tagRow($index = null, array $properties = null) {
        if ($this->sheetObject == null) {
            throw new \LogicException();
        }
        if ($index !== null && !is_int($index)) {
            throw new \InvalidArgumentException();
        }

        $this->column = null;
        $this->row = $index == null ? $this->increaseRow() : $index;

        $this->cellObject = null;

        if ($properties != null) {
            $this->setProperties($properties, $this->rowMappings);
        }
    }

    public function tagCell($index = null, $value = null, array $properties = null) {
        if ($this->sheetObject == null) {
            throw new \LogicException();
        }
        if ($index !== null && !is_int($index)) {
            throw new \InvalidArgumentException();
        }

        $this->column = $index == null ? $this->increaseColumn() : $index;

        $this->cellObject = $this->sheetObject->getCellByColumnAndRow($this->column, $this->row);

        if ($value !== null) {
            $this->cellObject->setValue($value);
        }
        
        if ($properties != null) {
            $this->setProperties($properties, $this->cellMappings);
        }
    }
    
    public function tagDrawing($path, array $properties = null) {
        if ($this->sheetObject == null) {
            throw new \LogicException();
        }
        if (!file_exists($path)) {
            throw new \InvalidArgumentException();
        }

        $this->drawingObject = new \PHPExcel_Worksheet_Drawing();
        $this->drawingObject->setWorksheet($this->sheetObject);
        $this->drawingObject->setPath($path);

        if ($properties != null) {
            $this->setProperties($properties, $this->cellMappings);
        }

        $this->drawingObject = null;
    }

    public function save($format = null) {
        if ($this->format != null) {
            $format = $this->format;
        }
        if ($format == null || empty($format)) {
            $format = 'xls';
        }

        $writerType = null;
        switch(strtolower($format)) {
            case 'csv':
                $writerType = 'CSV';
                break;
            case 'xls':
                $writerType = 'Excel5';
                break;
            case 'xlsx':
                $writerType = 'Excel2007';
                break;
            case 'pdf':
                $writerType = 'PDF';
                break;
            case 'default':
                throw new \InvalidArgumentException();
        }

        \PHPExcel_IOFactory::createWriter($this->documentObject, $writerType)->save('php://output');
    }

    private function increaseRow() {
        return $this->row === null ? self::$ROW_DEFAULT : $this->row + 1;
    }

    private function increaseColumn() {
        return $this->column === null ? self::$COLUMN_DEFAULT : $this->column + 1;
    }
}