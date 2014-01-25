<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Nodes;

class XlsCellNode extends \Twig_Node
{
    public function __construct(\Twig_Node_Expression $index, \Twig_Node_Expression $properties, \Twig_NodeInterface $body, $line, $tag = 'xlscell')
    {
        parent::__construct(['index' => $index, 'properties' => $properties, 'body' => $body], [], $line, $tag);
    }

    public function compile(\Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$cellIndex = ')
            ->subcompile($this->getNode('index'), true)
            ->raw(';'.PHP_EOL)

            ->write("ob_start();\n")
            ->subcompile($this->getNode('body'), true)
            ->write('$cellValue = trim(ob_get_clean());'.PHP_EOL)

            ->write('$cellProperties = ')
            ->subcompile($this->getNode('properties'), true)
            ->raw(';'.PHP_EOL)

            ->write('
                $phpExcel->tagCell($cellIndex, $cellValue, $cellProperties);
                unset($cellIndex, $cellValue, $cellProperties);
            ');
    }
}