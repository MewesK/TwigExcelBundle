<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;

/**
 * Class XlsSheetNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
class XlsSheetNode extends XlsNode
{
    /**
     * @param Twig_Node_Expression $index
     * @param Twig_Node_Expression $properties
     * @param Twig_Node $body
     * @param int $line
     * @param string $tag
     */
    public function __construct(Twig_Node_Expression $index, Twig_Node_Expression $properties, Twig_Node $body, $line = 0, $tag = 'xlssheet')
    {
        parent::__construct(['index' => $index, 'properties' => $properties, 'body' => $body], [], $line, $tag);
    }

    /**
     * @param Twig_Compiler $compiler
     */
    public function compile(Twig_Compiler $compiler)
    {
        $compiler->addDebugInfo($this)
            ->write('$sheetIndex = ')
            ->subcompile($this->getNode('index'))
            ->raw(';' . PHP_EOL)
            ->write('$sheetProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';' . PHP_EOL)
            ->write('$context[\'phpExcel\']->startSheet($sheetIndex, $sheetProperties);' . PHP_EOL)
            ->write('unset($sheetIndex, $sheetProperties);' . PHP_EOL);

        if ($this->hasNode('body')) {
            $compiler->subcompile($this->getNode('body'));
        }

        $compiler->addDebugInfo($this)->write('$context[\'phpExcel\']->endSheet();' . PHP_EOL);
    }

    /**
     * @return string[]
     */
    public function getAllowedParents()
    {
        return [
            'MewesK\TwigExcelBundle\Twig\Node\XlsDocumentNode'
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
