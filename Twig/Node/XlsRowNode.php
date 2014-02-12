<?php

namespace MewesK\TwigExcelBundle\Twig\Node;
use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;

class XlsRowNode extends Twig_Node
{
    public function __construct(Twig_Node_Expression $index, Twig_Node $body, $line, $tag = 'xlsrow')
    {
        parent::__construct(array('index' => $index, 'body' => $body), array(), $line, $tag);
    }

    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$rowIndex = ')
            ->subcompile($this->getNode('index'))
            ->raw(';'.PHP_EOL)

            ->write('$phpExcel->startRow($rowIndex);'.PHP_EOL)
            ->write('unset($rowIndex);'.PHP_EOL)

            ->subcompile($this->getNode('body'))

            ->addDebugInfo($this)

            ->write('$phpExcel->endRow();'.PHP_EOL);
    }
}