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
     * @var
     */
    protected $context;
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
    

    public function __construct($context) {
        $this->context = $context;

        $this->documentObject = new \PHPExcel();
        $this->documentObject->removeSheetByIndex(0);

        $this->sheetObject = null;
        $this->cellObject = null;
        $this->drawingObject = null;
        $this->row = null;
        $this->column = null;
        $this->format = null;

        $this->documentMappings = array();
        $this->sheetMappings = array();
        $this->rowMappings = array();
        $this->cellMappings = array();
        $this->drawingMappings = array();

        $this->initDocumentPropertyMappings();
        $this->initSheetPropertyMappings();
        $this->initRowPropertyMappings();
        $this->initCellPropertyMappings();
        $this->initDrawingPropertyMappings();
    }

    protected function initDocumentPropertyMappings() {
        $wrapper = $this; // PHP 5.3 fix

        $this->documentMappings['category'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setCategory($value); };
        $this->documentMappings['company'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setCompany($value); };
        $this->documentMappings['created'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setCreated($value); };
        $this->documentMappings['creator'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setCreator($value); };
        $this->documentMappings['defaultStyle'] = function($value) use ($wrapper) { $wrapper->documentObject->getDefaultStyle()->applyFromArray($value); };
        $this->documentMappings['description'] = function($value) use ($wrapper) { $wrapper->documentObject->getProperties()->setDescription($value); };
        $this->documentMappings['format'] = function($value) use ($wrapper) { $wrapper->format = $value; };
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

    protected function initSheetPropertyMappings() {
        $wrapper = $this; // PHP 5.3 fix

        $this->sheetMappings['columnDimension']['__multi'] = true;
        $this->sheetMappings['columnDimension']['__object'] = function($key) use ($wrapper) { return $key == 'default' ? $wrapper->sheetObject->getDefaultColumnDimension() : $wrapper->sheetObject->getColumnDimension($key); };
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
        $this->sheetMappings['protection']['pivotTables'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setPivotTables($value); };
        $this->sheetMappings['protection']['scenarios'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setScenarios($value); };
        $this->sheetMappings['protection']['selectLockedCells'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setSelectLockedCells($value); };
        $this->sheetMappings['protection']['selectUnlockedCells'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setSelectUnlockedCells($value); };
        $this->sheetMappings['protection']['sheet'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setSheet($value); };
        $this->sheetMappings['protection']['sort'] = function($value) use ($wrapper) { $wrapper->sheetObject->getProtection()->setSort($value); };
        $this->sheetMappings['rightToLeft'] = function($value) use ($wrapper) { $wrapper->sheetObject->setRightToLeft($value); };
        $this->sheetMappings['rowDimension']['__multi'] = true;
        $this->sheetMappings['rowDimension']['__object'] = function($key) use ($wrapper) { return $key == 'default' ? $wrapper->sheetObject->getDefaultRowDimension() : $wrapper->sheetObject->getRowDimension($key); };
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

    protected function initRowPropertyMappings() {
        // nothing
    }

    protected function initCellPropertyMappings() {
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

    protected function initDrawingPropertyMappings() {
        $wrapper = $this; // PHP 5.3 fix

        $this->drawingMappings['coordinates'] = function($value) use ($wrapper) { $wrapper->drawingObject->getCoordinates($value); };
        $this->drawingMappings['description'] = function($value) use ($wrapper) { $wrapper->drawingObject->getDescription($value); };
        $this->drawingMappings['height'] = function($value) use ($wrapper) { $wrapper->drawingObject->getHeight($value); };
        $this->drawingMappings['name'] = function($value) use ($wrapper) { $wrapper->drawingObject->getName($value); };
        $this->drawingMappings['offsetX'] = function($value) use ($wrapper) { $wrapper->drawingObject->getOffsetX($value); };
        $this->drawingMappings['offsetY'] = function($value) use ($wrapper) { $wrapper->drawingObject->getOffsetY($value); };
        $this->drawingMappings['resizeProportional'] = function($value) use ($wrapper) { $wrapper->drawingObject->getResizeProportional($value); };
        $this->drawingMappings['rotation'] = function($value) use ($wrapper) { $wrapper->drawingObject->getRotation($value); };
        $this->drawingMappings['shadow']['alignment'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setAlignment($value); };
        $this->drawingMappings['shadow']['alpha'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setAlpha($value); };
        $this->drawingMappings['shadow']['blurRadius'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setBlurRadius($value); };
        $this->drawingMappings['shadow']['color'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->getColor()->setRgb($value); };
        $this->drawingMappings['shadow']['direction'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setDirection($value); };
        $this->drawingMappings['shadow']['distance'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setDistance($value); };
        $this->drawingMappings['shadow']['visible'] = function($value) use ($wrapper) { $wrapper->drawingObject->getShadow()->setVisible($value); };
        $this->drawingMappings['width'] = function($value) use ($wrapper) { $wrapper->drawingObject->getWidth($value); };
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
        $this->row = null;
        $this->column = null;
        $this->format = null;

        $this->sheetObject = null;
        $this->cellObject = null;
        $this->drawingObject = null;

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

    public function tagHeaderFooter($type = null, $value = null, array $properties = null) {
        if ($this->sheetObject == null) {
            throw new \LogicException();
        }
        $headerObject = $this->sheetObject->getHeaderFooter();

        switch($type) {
            case 'header':
                $headerObject->setOddHeader($value);
                $headerObject->setEvenHeader($value);
                $headerObject->setFirstHeader($value);
                break;
            case 'footer':
                $headerObject->setOddFooter($value);
                $headerObject->setEvenFooter($value);
                $headerObject->setFirstFooter($value);
                break;
            case 'oddHeader':
                $headerObject->setDifferentOddEven(true);
                $headerObject->setOddHeader($value);
                break;
            case 'oddFooter':
                $headerObject->setDifferentOddEven(true);
                $headerObject->setOddFooter($value);
                break;
            case 'evenHeader':
                $headerObject->setDifferentOddEven(true);
                $headerObject->setEvenHeader($value);
                break;
            case 'evenFooter':
                $headerObject->setDifferentOddEven(true);
                $headerObject->setEvenFooter($value);
                break;
            case 'firstHeader':
                $headerObject->setDifferentFirst(true);
                $headerObject->setFirstHeader($value);
                break;
            case 'firstFooter':
                $headerObject->setDifferentFirst(true);
                $headerObject->setFirstFooter($value);
                break;
            default:
                throw new \InvalidArgumentException();
        }

        $this->sheetObject->setHeaderFooter($headerObject);

    }
    
    public function tagDrawing($path, array $properties = null) {
        try {
            $pathInfo = pathinfo($path, PATHINFO_EXTENSION);
            $tempPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'xlsdrawing' . '_' . md5($path) .
                (isset($pathInfo['extension']) && !empty($pathInfo['extension']) ? '.'.$pathInfo['extension'] : '');

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

            if ($this->sheetObject == null) {
                throw new \LogicException();
            }

            $this->drawingObject = new \PHPExcel_Worksheet_Drawing();
            $this->drawingObject->setWorksheet($this->sheetObject);
            $this->drawingObject->setPath($tempPath);

            if ($properties != null) {
                $this->setProperties($properties, $this->drawingMappings);
            }

            $this->drawingObject = null;
        } catch(\Exception $e) {
            $this->drawingObject = null;

            // re-throw
            throw $e;
        }
    }

    public function save() {
        $format = null;
        if ($this->format != null) {
            $format = $this->format;
        } else {
            $app = $this->context && isset($this->context["app"]) ? $this->context["app"] : null;
            $request = $app && isset($app['request']) ? $app['request'] : null;
            $format = $request && isset($request['requestFormat']) ? $request['requestFormat'] : null;
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