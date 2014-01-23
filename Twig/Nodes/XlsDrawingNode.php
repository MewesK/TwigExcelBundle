<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Nodes;

class XlsDrawingNode extends \Twig_Node
{
    public function __construct(\Twig_Node_Expression $path, \Twig_Node_Expression $properties, $lineno, $tag = 'xlsdrawing')
    {
        parent::__construct(['path' => $path, 'properties' => $properties], [], $lineno, $tag);
    }

    public function compile(\Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)
            ->write('$strPath = ')
            ->subcompile($this->getNode('path'), true)
            ->raw(';'.PHP_EOL)
            ->write('$arrDrawingProperties = ')
            ->subcompile($this->getNode('properties'), true)
            ->raw(';'.PHP_EOL)

            ->write('
                $objDrawing = new PHPExcel_Worksheet_Drawing();
                $objDrawing->setWorksheet($objActiveSheet);
                $objDrawing->setPath($strPath);

                if (count($arrDrawingProperties) > 0) {
                    if (array_key_exists(\'coordinates\', $arrDrawingProperties)) { $objDrawing->getCoordinates($arrDrawingProperties[\'coordinates\']); }
                    if (array_key_exists(\'description\', $arrDrawingProperties)) { $objDrawing->getDescription($arrDrawingProperties[\'description\']); }
                    if (array_key_exists(\'height\', $arrDrawingProperties)) { $objDrawing->getHeight($arrDrawingProperties[\'height\']); }
                    if (array_key_exists(\'name\', $arrDrawingProperties)) { $objDrawing->getName($arrDrawingProperties[\'name\']); }
                    if (array_key_exists(\'offsetX\', $arrDrawingProperties)) { $objDrawing->getOffsetX($arrDrawingProperties[\'offsetX\']); }
                    if (array_key_exists(\'offsetY\', $arrDrawingProperties)) { $objDrawing->getOffsetY($arrDrawingProperties[\'offsetY\']); }
                    if (array_key_exists(\'resizeProportional\', $arrDrawingProperties)) { $objDrawing->getResizeProportional($arrDrawingProperties[\'resizeProportional\']); }
                    if (array_key_exists(\'rotation\', $arrDrawingProperties)) { $objDrawing->getRotation($arrDrawingProperties[\'rotation\']); }

                    if (array_key_exists(\'shadow\', $arrDrawingProperties)) {
                        $arrShadow = $arrDrawingProperties[\'shadow\'];
                        $objShadow = $objDrawing->getShadow();

                        if (array_key_exists(\'alignment\', $arrShadow)) { $objShadow->setAlignment($arrShadow[\'alignment\']); }
                        if (array_key_exists(\'alpha\', $arrShadow)) { $objShadow->setAlpha($arrShadow[\'alpha\']); }
                        if (array_key_exists(\'blurRadius\', $arrShadow)) { $objShadow->setBlurRadius($arrShadow[\'blurRadius\']); }
                        if (array_key_exists(\'color\', $arrShadow)) { $objShadow->setColor()->setRgb($arrShadow[\'color\']); }
                        if (array_key_exists(\'direction\', $arrShadow)) { $objShadow->setDirection($arrShadow[\'direction\']); }
                        if (array_key_exists(\'distance\', $arrShadow)) { $objShadow->setDistance($arrShadow[\'distance\']); }
                        if (array_key_exists(\'visible\', $arrShadow)) { $objShadow->setVisible($arrShadow[\'visible\']); }

                        unset($arrShadow, $objShadow);
                    }

                    if (array_key_exists(\'width\', $arrDrawingProperties)) { $objDrawing->getWidth($arrDrawingProperties[\'width\']); }
                }

                unset($objDrawing, $strPath, $arrDrawingProperties);
            ');
    }
}