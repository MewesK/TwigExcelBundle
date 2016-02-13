<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;

/**
 * Class XlsRightNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
class XlsRightNode extends XlsNode
{
    /**
     * @param Twig_Node $body
     * @param int $line
     * @param string $tag
     */
    public function __construct(Twig_Node $body, $line = 0, $tag = 'xlsright')
    {
        parent::__construct(['body' => $body], [], $line, $tag);
    }

    /**
     * @param Twig_Compiler $compiler
     */
    public function compile(Twig_Compiler $compiler)
    {
        $compiler->addDebugInfo($this)
            ->write('$context[\'phpExcel\']->startAlignment(\'right\');' . PHP_EOL)
            ->write("ob_start();\n")
            ->subcompile($this->getNode('body'))
            ->write('$rightValue = trim(ob_get_clean());' . PHP_EOL)
            ->write('$context[\'phpExcel\']->endAlignment($rightValue);' . PHP_EOL)
            ->write('unset($rightValue);' . PHP_EOL);
    }

    /**
     * @return string[]
     */
    public function getAllowedParents()
    {
        return [
            'MewesK\TwigExcelBundle\Twig\Node\XlsFooterNode',
            'MewesK\TwigExcelBundle\Twig\Node\XlsHeaderNode'
        ];
    }

    /**
     * @return bool
     */
    public function canContainText()
    {
        return true;
    }
}
