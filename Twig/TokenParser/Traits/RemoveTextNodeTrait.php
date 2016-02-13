<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser\Traits;

use MewesK\TwigExcelBundle\Twig\Node\XlsCellNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsCenterNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsLeftNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsRightNode;
use Twig_Node;
use Twig_Node_Block;
use Twig_Node_BlockReference;
use Twig_Node_Text;

/**
 * Class RemoveTextNodeTrait
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
trait RemoveTextNodeTrait
{
    /**
     * @param Twig_Node $node
     */
    protected function removeTextNodesRecursively(Twig_Node $node)
    {
        foreach ($node->getIterator() as $key => $subNode) {
            if ($subNode instanceof Twig_Node_Text) {
                // Never delete a block body
                if ($key === 'body' && $node instanceof Twig_Node_Block) {
                    continue;
                }

                $node->removeNode($key);
            } elseif ($subNode instanceof Twig_Node_BlockReference) {
                $this->removeTextNodesRecursively($this->parser->getBlock($subNode->getAttribute('name')));
            } elseif ($subNode instanceof Twig_Node && $subNode->count() > 0) {
                if ($subNode instanceof XlsNode && $subNode->canContainText()) {
                    continue;
                }
                $this->removeTextNodesRecursively($subNode);
            }
        }
    }
}
