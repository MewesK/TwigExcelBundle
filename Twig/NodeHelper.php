<?php

namespace MewesK\TwigExcelBundle\Twig;

use MewesK\TwigExcelBundle\Twig\Node\XlsNode;
use Twig_Error_Syntax;
use Twig_Node;
use Twig_Node_Block;
use Twig_Node_BlockReference;
use Twig_Node_Expression_MethodCall;
use Twig_Node_Expression_Name;
use Twig_Node_Text;
use Twig_Parser;

/**
 * Class NodeHelper
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class NodeHelper
{
    /**
     * Adds the PhpExcel instance as the last parameter to all macro function calls.
     *
     * @param Twig_Node $node
     */
    public static function fixMacroCallsRecursively(Twig_Node $node)
    {
        foreach ($node->getIterator() as $key => $subNode) {
            if ($subNode instanceof Twig_Node_Expression_MethodCall) {
                /**
                 * @var \Twig_Node_Expression_Array $argumentsNode
                 */
                $argumentsNode = $subNode->getNode('arguments');
                $argumentsNode->addElement(new Twig_Node_Expression_Name('phpExcel', null), null);
            } elseif ($subNode instanceof Twig_Node && $subNode->count() > 0) {
                self::fixMacroCallsRecursively($subNode);
            }
        }
    }

    /**
     * Removes all TextNodes that are in illegal places to avoid echos in unwanted places.
     *
     * @param Twig_Node $node
     * @param Twig_Parser $parser
     */
    public static function removeTextNodesRecursively(Twig_Node $node, Twig_Parser $parser)
    {
        foreach ($node->getIterator() as $key => $subNode) {
            if ($subNode instanceof Twig_Node_Text) {
                // Never delete a block body
                if ($key === 'body' && $node instanceof Twig_Node_Block) {
                    continue;
                }

                $node->removeNode($key);
            } elseif ($subNode instanceof Twig_Node_BlockReference) {
                self::removeTextNodesRecursively($parser->getBlock($subNode->getAttribute('name')), $parser);
            } elseif ($subNode instanceof Twig_Node && $subNode->count() > 0) {
                if ($subNode instanceof XlsNode && $subNode->canContainText()) {
                    continue;
                }
                self::removeTextNodesRecursively($subNode, $parser);
            }
        }
    }

    /**
     * @param XlsNode $node
     * @param array $path
     * @throws Twig_Error_Syntax
     */
    public static function checkAllowedParents(XlsNode $node, array $path)
    {
        $parentName = null;

        foreach (array_reverse($path) as $className) {
            if (strpos($className, 'MewesK\TwigExcelBundle\Twig\Node\Xls') === 0) {
                $parentName = $className;
                break;
            }
        }

        if ($parentName === null) {
            return;
        }

        foreach ($node->getAllowedParents() as $className) {
            if ($className === $parentName) {
                return;
            }
        }

        throw new Twig_Error_Syntax(sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), $parentName));
    }

    /**
     * @param Twig_Node $node
     * @return bool
     */
    public static function checkContainsXlsNode(Twig_Node $node)
    {
        foreach ($node->getIterator() as $key => $subNode) {
            if ($node instanceof XlsNode) {
                return true;
            } elseif ($subNode instanceof Twig_Node && $subNode->count() > 0) {
                if (self::checkContainsXlsNode($subNode)) {
                    return true;
                }
            }
        }
        return false;
    }
}
