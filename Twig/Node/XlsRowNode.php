<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Node;
use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;
use Twig_NodeInterface;

class XlsRowNode extends Twig_Node
{
    public function __construct(Twig_Node_Expression $index, Twig_Node_Expression $properties, Twig_NodeInterface $body, $line, $tag = 'xlsrow')
    {
        parent::__construct(array('index' => $index, 'properties' => $properties, 'body' => $body), array(), $line, $tag);
    }

    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$rowIndex = ')
            ->subcompile($this->getNode('index'))
            ->raw(';'.PHP_EOL)

            ->write('$rowProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';'.PHP_EOL)

            ->write('$phpExcel->startRow($rowIndex, $rowProperties);'.PHP_EOL)
            ->write('unset($rowIndex, $rowProperties);'.PHP_EOL)

            ->subcompile($this->getNode('body'))

            ->addDebugInfo($this)

            ->write('$phpExcel->endRow();'.PHP_EOL);
    }
}