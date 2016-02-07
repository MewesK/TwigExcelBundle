<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use Twig_Node;
use Twig_Node_Expression_MethodCall;
use Twig_Node_Expression_Name;

/**
 * Class FixMacroCallsTrait
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
trait FixMacroCallsTrait
{
    /**
     * @param Twig_Node $node
     */
    private function fixMacroCallsRecursively(Twig_Node $node)
    {
        foreach ($node->getIterator() as $key => $subNode) {
            if ($subNode instanceof Twig_Node_Expression_MethodCall) {
                /**
                 * @var \Twig_Node_Expression_Array $argumentsNode
                 */
                $argumentsNode = $subNode->getNode('arguments');
                $argumentsNode->addElement(new Twig_Node_Expression_Name('phpExcel', null), null);
            } elseif ($subNode instanceof Twig_Node && $subNode->count() > 0) {
                $this->fixMacroCallsRecursively($subNode);
            }
        }
    }
}
