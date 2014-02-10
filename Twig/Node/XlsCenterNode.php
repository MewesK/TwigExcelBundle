<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;
use Twig_NodeInterface;

class XlsCenterNode extends Twig_Node
{
    public function __construct(Twig_NodeInterface $body, $line, $tag = 'xlscenter')
    {
        parent::__construct(array('body' => $body), array(), $line, $tag);
    }

    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$phpExcel->startAlignment(\'center\');'.PHP_EOL)

            ->write("ob_start();\n")
            ->subcompile($this->getNode('body'))
            ->write('$centerValue = trim(ob_get_clean());'.PHP_EOL)

            ->write('$phpExcel->endAlignment($centerValue);'.PHP_EOL)
            ->write('unset($centerValue);'.PHP_EOL);
    }
}