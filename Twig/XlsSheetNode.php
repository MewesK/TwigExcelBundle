<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig;

class XlsSheetNode extends \Twig_Node
{
    public function __construct(\Twig_Node_Expression $title, \Twig_Node_Expression $properties, \Twig_NodeInterface $body, $lineno, $tag = 'xlssheet')
    {
        parent::__construct(['title' => $title, 'properties' => $properties, 'body' => $body], [], $lineno, $tag);
    }

    public function compile(\Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)
            ->write('$strTitle = ')
            ->subcompile($this->getNode('title'), true)
            ->raw(';'.PHP_EOL)
            ->write('$arrSheetProperties = ')
            ->subcompile($this->getNode('properties'), true)
            ->raw(';'.PHP_EOL)

            ->write('
                if (!$objPHPExcel->sheetNameExists($strTitle)) { $objPHPExcel->createSheet()->setTitle($strTitle); }
                $objPHPExcel->setActiveSheetIndexByName($strTitle);
                $objActiveSheet = $objPHPExcel->getActiveSheet();

                if (count($arrSheetProperties) > 0) {
                    if (array_key_exists(\'columnDimension\', $arrSheetProperties)) {
                        $arrColumnDimensions = $arrSheetProperties[\'columnDimension\'];

                        foreach ($arrColumnDimensions as $arrColumn => $arrColumnDimension) {
                            $objColumnDimension = $arrColumn == \'default\' ? $objActiveSheet->getDefaultColumnDimension() : $objActiveSheet->getColumnDimension($arrColumn);

                            if (array_key_exists(\'autoSize\', $arrColumnDimension)) { $objColumnDimension->setAutoSize($arrColumnDimension[\'autoSize\']); }
                            if (array_key_exists(\'collapsed\', $arrColumnDimension)) { $objColumnDimension->setCollapsed($arrColumnDimension[\'collapsed\']); }
                            if (array_key_exists(\'columnIndex\', $arrColumnDimension)) { $objColumnDimension->setColumnIndex($arrColumnDimension[\'columnIndex\']); }
                            if (array_key_exists(\'outlineLevel\', $arrColumnDimension)) { $objColumnDimension->setOutlineLevel($arrColumnDimension[\'outlineLevel\']); }
                            if (array_key_exists(\'visible\', $arrColumnDimension)) { $objColumnDimension->setVisible($arrColumnDimension[\'visible\']); }
                            if (array_key_exists(\'width\', $arrColumnDimension)) { $objColumnDimension->setWidth($arrColumnDimension[\'width\']); }
                            if (array_key_exists(\'xfIndex\', $arrColumnDimension)) { $objColumnDimension->setXfIndex($arrColumnDimension[\'xfIndex\']); }

                            unset($objColumnDimension);
                        }

                        unset($arrColumnDimensions);
                    }

                    if (array_key_exists(\'columnDimension\', $arrSheetProperties)) {
                        $arrColumnDimensions = $arrSheetProperties[\'columnDimension\'];

                        foreach ($arrColumnDimensions as $arrColumn => $arrColumnDimension) {
                            $objColumnDimension = $objActiveSheet->getColumnDimension($arrColumn);

                            if (array_key_exists(\'autoSize\', $arrColumnDimension)) { $objColumnDimension->setAutoSize($arrColumnDimension[\'autoSize\']); }
                            if (array_key_exists(\'collapsed\', $arrColumnDimension)) { $objColumnDimension->setCollapsed($arrColumnDimension[\'collapsed\']); }
                            if (array_key_exists(\'columnIndex\', $arrColumnDimension)) { $objColumnDimension->setColumnIndex($arrColumnDimension[\'columnIndex\']); }
                            if (array_key_exists(\'outlineLevel\', $arrColumnDimension)) { $objColumnDimension->setOutlineLevel($arrColumnDimension[\'outlineLevel\']); }
                            if (array_key_exists(\'visible\', $arrColumnDimension)) { $objColumnDimension->setVisible($arrColumnDimension[\'visible\']); }
                            if (array_key_exists(\'width\', $arrColumnDimension)) { $objColumnDimension->setWidth($arrColumnDimension[\'width\']); }
                            if (array_key_exists(\'xfIndex\', $arrColumnDimension)) { $objColumnDimension->setXfIndex($arrColumnDimension[\'xfIndex\']); }

                            unset($objColumnDimension);
                        }

                        unset($arrColumnDimensions);
                    }

                    if (array_key_exists(\'footer\', $arrSheetProperties)) { $objActiveSheet->getHeaderFooter()->setOddFooter($arrSheetProperties[\'footer\']); }
                    if (array_key_exists(\'header\', $arrSheetProperties)) { $objActiveSheet->getHeaderFooter()->setOddHeader($arrSheetProperties[\'header\']); }

                    if (array_key_exists(\'pageMargins\', $arrSheetProperties)) {
                        $arrPageMargins = $arrSheetProperties[\'pageMargins\'];
                        $objPageMargins = $objActiveSheet->getPageMargins();

                        if (array_key_exists(\'top\', $arrPageMargins)) { $objPageMargins->setTop($arrPageMargins[\'top\']); }
                        if (array_key_exists(\'bottom\', $arrPageMargins)) { $objPageMargins->setBottom($arrPageMargins[\'bottom\']); }
                        if (array_key_exists(\'left\', $arrPageMargins)) { $objPageMargins->setLeft($arrPageMargins[\'left\']); }
                        if (array_key_exists(\'right\', $arrPageMargins)) { $objPageMargins->setRight($arrPageMargins[\'right\']); }
                        if (array_key_exists(\'header\', $arrPageMargins)) { $objPageMargins->setHeader($arrPageMargins[\'header\']); }
                        if (array_key_exists(\'footer\', $arrPageMargins)) { $objPageMargins->setFooter($arrPageMargins[\'footer\']); }

                        unset($arrPageMargins, $objPageMargins);
                    }

                    if (array_key_exists(\'pageSetup\', $arrSheetProperties)) {
                        $arrPageSetup = $arrSheetProperties[\'pageSetup\'];
                        $objPageSetup = $objActiveSheet->getPageSetup();

                        if (array_key_exists(\'fitToHeight\', $arrPageSetup)) { $objPageSetup->setFitToHeight($arrPageSetup[\'fitToHeight\']); }
                        if (array_key_exists(\'fitToPage\', $arrPageSetup)) { $objPageSetup->setFitToPage($arrPageSetup[\'fitToPage\']); }
                        if (array_key_exists(\'fitToWidth\', $arrPageSetup)) { $objPageSetup->setFitToWidth($arrPageSetup[\'fitToWidth\']); }
                        if (array_key_exists(\'horizontalCentered\', $arrPageSetup)) { $objPageSetup->setHorizontalCentered($arrPageSetup[\'horizontalCentered\']); }
                        if (array_key_exists(\'orientation\', $arrPageSetup)) { $objPageSetup->setOrientation($arrPageSetup[\'orientation\']); }
                        if (array_key_exists(\'paperSize\', $arrPageSetup)) { $objPageSetup->setPaperSize($arrPageSetup[\'paperSize\']); }
                        if (array_key_exists(\'printArea\', $arrPageSetup)) { $objPageSetup->setPrintArea($arrPageSetup[\'printArea\']); }
                        if (array_key_exists(\'scale\', $arrPageSetup)) { $objPageSetup->setScale($arrPageSetup[\'scale\']); }
                        if (array_key_exists(\'verticalCentered\', $arrPageSetup)) { $objPageSetup->setVerticalCentered($arrPageSetup[\'verticalCentered\']); }

                        unset($arrPageSetup, $objPageSetup);
                    }

                    if (array_key_exists(\'protection\', $arrSheetProperties)) {
                        $arrProtection = $arrSheetProperties[\'protection\'];
                        $objProtection = $objActiveSheet->getProtection();

                        if (array_key_exists(\'autoFilter\', $arrProtection)) { $objProtection->setAutoFilter($arrPageSetup[\'autoFilter\']); }
                        if (array_key_exists(\'deleteColumns\', $arrProtection)) { $objProtection->setDeleteColumns($arrPageSetup[\'deleteColumns\']); }
                        if (array_key_exists(\'deleteRows\', $arrProtection)) { $objProtection->setDeleteRows($arrPageSetup[\'deleteRows\']); }
                        if (array_key_exists(\'formatCells\', $arrProtection)) { $objProtection->setFormatCells($arrPageSetup[\'formatCells\']); }
                        if (array_key_exists(\'formatColumns\', $arrProtection)) { $objProtection->setFormatColumns($arrPageSetup[\'formatColumns\']); }
                        if (array_key_exists(\'formatRows\', $arrProtection)) { $objProtection->setFormatRows($arrPageSetup[\'formatRows\']); }
                        if (array_key_exists(\'insertColumns\', $arrProtection)) { $objProtection->setInsertColumns($arrPageSetup[\'insertColumns\']); }
                        if (array_key_exists(\'insertHyperlinks\', $arrProtection)) { $objProtection->setInsertHyperlinks($arrPageSetup[\'insertHyperlinks\']); }
                        if (array_key_exists(\'insertRows\', $arrProtection)) { $objProtection->setInsertRows($arrPageSetup[\'insertRows\']); }
                        if (array_key_exists(\'objects\', $arrProtection)) { $objProtection->setObjects($arrPageSetup[\'objects\']); }
                        if (array_key_exists(\'pivotTables\', $arrProtection)) { $objProtection->setPivotTables($arrPageSetup[\'pivotTables\']); }
                        if (array_key_exists(\'scenarios\', $arrProtection)) { $objProtection->setSelectLockedCells($arrPageSetup[\'scenarios\']); }
                        if (array_key_exists(\'selectLockedCells\', $arrProtection)) { $objProtection->setVerticalCentered($arrPageSetup[\'selectLockedCells\']); }
                        if (array_key_exists(\'selectUnlockedCells\', $arrProtection)) { $objProtection->setSelectUnlockedCells($arrPageSetup[\'selectUnlockedCells\']); }
                        if (array_key_exists(\'sheet\', $arrProtection)) { $objProtection->setSheet($arrPageSetup[\'sheet\']); }
                        if (array_key_exists(\'sort\', $arrProtection)) { $objProtection->setSort($arrPageSetup[\'sort\']); }

                        unset($arrProtection, $objProtection);
                    }

                    if (array_key_exists(\'printGridlines\', $arrSheetProperties)) { $objActiveSheet->setPrintGridlines($arrSheetProperties[\'printGridlines\']); }
                    if (array_key_exists(\'rightToLeft\', $arrSheetProperties)) { $objActiveSheet->setRightToLeft($arrSheetProperties[\'rightToLeft\']); }

                    if (array_key_exists(\'rowDimension\', $arrSheetProperties)) {
                        $arrRowDimensions = $arrSheetProperties[\'rowDimension\'];

                        foreach ($arrRowDimensions as $arrRow => $arrRowDimension) {
                            $objRowDimension = $arrRow == \'default\' ? $objActiveSheet->getDefaultRowDimension() : $objActiveSheet->getRowDimension($arrRow);

                            if (array_key_exists(\'collapsed\', $arrRowDimension)) { $objRowDimension->setCollapsed($arrRowDimension[\'collapsed\']); }
                            if (array_key_exists(\'outlineLevel\', $arrRowDimension)) { $objRowDimension->setOutlineLevel($arrRowDimension[\'outlineLevel\']); }
                            if (array_key_exists(\'rowHeight\', $arrRowDimension)) { $objRowDimension->setRowHeight($arrRowDimension[\'rowHeight\']); }
                            if (array_key_exists(\'rowIndex\', $arrRowDimension)) { $objRowDimension->setRowIndex($arrRowDimension[\'rowIndex\']); }
                            if (array_key_exists(\'visible\', $arrRowDimension)) { $objRowDimension->setVisible($arrRowDimension[\'visible\']); }
                            if (array_key_exists(\'xfIndex\', $arrRowDimension)) { $objRowDimension->setXfIndex($arrRowDimension[\'xfIndex\']); }
                            if (array_key_exists(\'zeroHeight\', $arrRowDimension)) { $objRowDimension->setZeroHeight($arrRowDimension[\'zeroHeight\']); }

                            unset($objRowDimension);
                        }

                        unset($arrRowDimensions);
                    }

                    if (array_key_exists(\'sheetState\', $arrSheetProperties)) { $objActiveSheet->setSheetState($arrSheetProperties[\'sheetState\']); }
                    if (array_key_exists(\'showGridlines\', $arrSheetProperties)) { $objActiveSheet->setShowGridlines($arrSheetProperties[\'showGridlines\']); }
                    if (array_key_exists(\'tabColor\', $arrSheetProperties)) { $objActiveSheet->getTabColor()->setRGB($arrSheetProperties[\'tabColor\']); }
                    if (array_key_exists(\'zoomScale\', $arrSheetProperties)) { $objActiveSheet->getSheetView()->setZoomScale($arrSheetProperties[\'zoomScale\']); }
                }

                unset($strTitle, $arrSheetProperties);
            ')

            ->subcompile($this->getNode('body'), true)
            ->write('unset($objActiveSheet);'.PHP_EOL);
    }
}