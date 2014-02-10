<?php

namespace MewesK\TwigExcelBundle\Twig\Node;
use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;
use Twig_NodeInterface;

class XlsFooterNode extends Twig_Node
{
    public function __construct(Twig_Node_Expression $type, Twig_Node_Expression $properties, Twig_NodeInterface $body, $line, $tag = 'xlsfooter')
    {
        parent::__construct(array('type' => $type, 'properties' => $properties, 'body' => $body), array(), $line, $tag);
    }

    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$footerType = ')
            ->subcompile($this->getNode('type'))
            ->raw(';'.PHP_EOL)
            ->write('$footerType = $footerType ? $footerType : \'footer\';'.PHP_EOL)

            ->write('$footerProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';'.PHP_EOL)

            ->write('$phpExcel->startHeaderFooter($footerType, $footerProperties);'.PHP_EOL)
            ->write('unset($footerType, $footerProperties);'.PHP_EOL)

            ->subcompile($this->getNode('body'))

            ->addDebugInfo($this)

            ->write('$phpExcel->endHeaderFooter();'.PHP_EOL);
    }
}