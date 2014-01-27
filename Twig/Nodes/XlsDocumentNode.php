<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Nodes;

class XlsDocumentNode extends \Twig_Node
{
    public function __construct(\Twig_Node_Expression $properties, \Twig_NodeInterface $body, $line, $tag = 'xlsdocument')
    {
        parent::__construct(['properties' => $properties, 'body' => $body], [], $line, $tag);
    }

    public function compile(\Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$documentProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';'.PHP_EOL)

            ->write('
                $phpExcel = new MewesK\PhpExcelTwigExtensionBundle\Twig\PhpExcelWrapper();
                $phpExcel->tagDocument($documentProperties);
                unset($documentProperties);
            ')

            ->subcompile($this->getNode('body'))

            ->write('
                $phpExcel->save($this->getAttribute($this->getAttribute((isset($context["app"]) ? $context["app"] : $this->getContext($context, "app")), "request"), "requestFormat"));
                unset($phpExcel);
            ');
    }
}