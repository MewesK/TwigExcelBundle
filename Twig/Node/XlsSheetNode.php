<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Node;
use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;
use Twig_NodeInterface;

class XlsSheetNode extends Twig_Node
{
    public function __construct(Twig_Node_Expression $index, Twig_Node_Expression $properties, Twig_NodeInterface $body, $line, $tag = 'xlssheet')
    {
        parent::__construct(array('index' => $index, 'properties' => $properties, 'body' => $body), array(), $line, $tag);
    }

    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$sheetIndex = ')
            ->subcompile($this->getNode('index'))
            ->raw(';'.PHP_EOL)

            ->write('$sheetProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';'.PHP_EOL)

            ->write('$phpExcel->tagSheet($sheetIndex, $sheetProperties);'.PHP_EOL)
            ->write('unset($sheetIndex, $sheetProperties);'.PHP_EOL)

            ->subcompile($this->getNode('body'));
    }
}