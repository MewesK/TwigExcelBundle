<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use MewesK\TwigExcelBundle\Twig\Node\XlsCellNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsCenterNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsDocumentNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsLeftNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsRightNode;
use Twig_Node;
use Twig_Node_Text;
use Twig_Token;

/**
 * Class XlsDocumentTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class XlsDocumentTokenParser extends AbstractTokenParser
{
    /**
     * @param Twig_Token $token
     *
     * @return XlsDocumentNode
     * @throws \Twig_Error_Syntax
     */
    public function parse(Twig_Token $token)
    {
        // parse attributes
        $properties = $this->parseProperties($token);
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);

        // parse body
        $body = $this->parseBody();
        $this->removeTextNodesRecursively($body);

        // return node
        return new XlsDocumentNode($properties, $body, $token->getLine(), $this->getTag());
    }

    /**
     * @return string
     */
    public function getTag()
    {
        return 'xlsdocument';
    }

    /**
     * @param Twig_Node $node
     */
    private function removeTextNodesRecursively(Twig_Node &$node)
    {
        foreach ($node->getIterator() as $key => $subNode) {
            if ($subNode instanceof Twig_Node_Text) {
                $node->removeNode($key);
            } elseif ($subNode instanceof Twig_Node && !($subNode instanceof XlsCellNode || $subNode instanceof XlsLeftNode || $subNode instanceof XlsCenterNode || $subNode instanceof XlsRightNode) && $subNode->count() > 0) {
                $this->removeTextNodesRecursively($subNode);
            }
        }
    }
}
