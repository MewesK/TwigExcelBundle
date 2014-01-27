<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Nodes;

class XlsRowNode extends \Twig_Node
{
    public function __construct(\Twig_Node_Expression $index, \Twig_Node_Expression $properties, \Twig_NodeInterface $body, $line, $tag = 'xlsrow')
    {
        parent::__construct(['index' => $index, 'properties' => $properties, 'body' => $body], [], $line, $tag);
    }

    public function compile(\Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$rowIndex = ')
            ->subcompile($this->getNode('index'))
            ->raw(';'.PHP_EOL)

            ->write('$rowProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';'.PHP_EOL)

            ->write('
                $phpExcel->tagRow($rowIndex, $rowProperties);
                unset($rowIndex, $rowProperties);
            ')

            ->subcompile($this->getNode('body'));
    }
}