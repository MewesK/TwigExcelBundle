<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig;

class XlsStyleNode extends \Twig_Node
{
    public function __construct(\Twig_Node_Expression $coordinates, \Twig_Node_Expression $properties, $lineno, $tag = 'xlsstyle')
    {
        parent::__construct(['coordinates' => $coordinates, 'properties' => $properties], [], $lineno, $tag);
    }

    public function compile(\Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)
            ->write('$strCoordinates = ')
            ->subcompile($this->getNode('coordinates'), true)
            ->raw(';'.PHP_EOL)
            ->write('$arrCellProperties = ')
            ->subcompile($this->getNode('properties'), true)
            ->raw(';'.PHP_EOL)

            ->write('
                if (array_key_exists(\'style\', $arrCellProperties)) { $objActiveSheet->getStyle($strCoordinates)->applyFromArray($arrCellProperties[\'style\']); }

                unset($strCoordinates, $arrCellProperties);
            ');
    }
}