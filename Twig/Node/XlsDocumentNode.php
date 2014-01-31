<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Node;
use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;
use Twig_NodeInterface;

class XlsDocumentNode extends Twig_Node
{
    public function __construct(Twig_Node_Expression $properties, Twig_NodeInterface $body, $line, $tag = 'xlsdocument')
    {
        parent::__construct(['properties' => $properties, 'body' => $body], [], $line, $tag);
    }

    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$documentProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';'.PHP_EOL)

            ->write('$phpExcel = new MewesK\PhpExcelTwigExtensionBundle\Twig\PhpExcelWrapper($context);'.PHP_EOL)
            ->write('$phpExcel->tagDocument($documentProperties);'.PHP_EOL)
            ->write('unset($documentProperties);'.PHP_EOL)

            ->subcompile($this->getNode('body'))

            ->write('$phpExcel->save();'.PHP_EOL)
            ->write('unset($phpExcel);'.PHP_EOL);

    }
}