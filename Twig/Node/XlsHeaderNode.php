<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;

/**
 * Class XlsHeaderNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
class XlsHeaderNode extends XlsNode
{
    /**
     * @param Twig_Node_Expression $type
     * @param Twig_Node_Expression $properties
     * @param Twig_Node $body
     * @param int $line
     * @param string $tag
     */
    public function __construct(Twig_Node_Expression $type, Twig_Node_Expression $properties, Twig_Node $body, $line = 0, $tag = 'xlsheader')
    {
        parent::__construct(['type' => $type, 'properties' => $properties, 'body' => $body], [], $line, $tag);
    }

    /**
     * @param Twig_Compiler $compiler
     */
    public function compile(Twig_Compiler $compiler)
    {
        $compiler->addDebugInfo($this)
            ->write('$headerType = ')
            ->subcompile($this->getNode('type'))
            ->raw(';' . PHP_EOL)
            ->write('$headerType = $headerType ? $headerType : \'header\';' . PHP_EOL)
            ->write('$headerProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';' . PHP_EOL)
            ->write('$context[\'phpExcel\']->startHeaderFooter($headerType, $headerProperties);' . PHP_EOL)
            ->write('unset($headerType, $headerProperties);' . PHP_EOL)
            ->subcompile($this->getNode('body'))
            ->addDebugInfo($this)
            ->write('$context[\'phpExcel\']->endHeaderFooter();' . PHP_EOL);
    }

    /**
     * @return string[]
     */
    public function getAllowedParents()
    {
        return [
            'MewesK\TwigExcelBundle\Twig\Node\XlsSheetNode'
        ];
    }

    /**
     * @return bool
     */
    public function canContainText()
    {
        return false;
    }
}
