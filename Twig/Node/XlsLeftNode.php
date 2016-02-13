<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;

/**
 * Class XlsLeftNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
class XlsLeftNode extends XlsNode
{
    /**
     * @param Twig_Node $body
     * @param int $line
     * @param string $tag
     */
    public function __construct(Twig_Node $body, $line = 0, $tag = 'xlsleft')
    {
        parent::__construct(['body' => $body], [], $line, $tag);
    }

    /**
     * @param Twig_Compiler $compiler
     */
    public function compile(Twig_Compiler $compiler)
    {
        $compiler->addDebugInfo($this)
            ->write('$context[\'phpExcel\']->startAlignment(\'left\');' . PHP_EOL)
            ->write("ob_start();\n")
            ->subcompile($this->getNode('body'))
            ->write('$leftValue = trim(ob_get_clean());' . PHP_EOL)
            ->write('$context[\'phpExcel\']->endAlignment($leftValue);' . PHP_EOL)
            ->write('unset($leftValue);' . PHP_EOL);
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
