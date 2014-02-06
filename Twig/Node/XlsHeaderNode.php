<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Node;
use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;
use Twig_NodeInterface;

class XlsHeaderNode extends Twig_Node
{
    public function __construct(Twig_Node_Expression $type, Twig_Node_Expression $properties, Twig_NodeInterface $body, $line, $tag = 'xlsheader')
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
            ->write('unset($headerProperties);'.PHP_EOL)

            ->write("ob_start();\n")
            ->subcompile($this->getNode('body'))
            ->write('$headerValue = trim(ob_get_clean());'.PHP_EOL)

            ->write('$phpExcel->endHeaderFooter($headerType, $headerValue);'.PHP_EOL)
            ->write('unset($headerType, $headerValue);'.PHP_EOL);
    }
}