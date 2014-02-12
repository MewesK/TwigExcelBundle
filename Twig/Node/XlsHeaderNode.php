<?php

namespace MewesK\TwigExcelBundle\Twig\Node;
use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;

class XlsHeaderNode extends Twig_Node
{
    public function __construct(Twig_Node_Expression $type, Twig_Node_Expression $properties, Twig_Node $body, $line, $tag = 'xlsheader')
    {
        parent::__construct(array('type' => $type, 'properties' => $properties, 'body' => $body), array(), $line, $tag);
    }

    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$headerType = ')
            ->subcompile($this->getNode('type'))
            ->raw(';'.PHP_EOL)
            ->write('$headerType = $headerType ? $headerType : \'header\';'.PHP_EOL)

            ->write('$headerProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';'.PHP_EOL)

            ->write('$phpExcel->startHeaderFooter($headerType, $headerProperties);'.PHP_EOL)
            ->write('unset($headerType, $headerProperties);'.PHP_EOL)

            ->subcompile($this->getNode('body'))

            ->addDebugInfo($this)

            ->write('$phpExcel->endHeaderFooter();'.PHP_EOL);
    }
}