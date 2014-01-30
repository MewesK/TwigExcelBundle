<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\TokenParser;

use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsCellNode;
use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsDocumentNode;
use Twig_Node;
use Twig_Node_Expression_Array;
use Twig_Node_Text;
use Twig_Token;
use Twig_TokenParser;

class XlsDocumentTokenParser extends Twig_TokenParser
{
    public function parse(Twig_Token $token)
    {
        $properties = new Twig_Node_Expression_Array([], $token->getLine());
        if (!$this->parser->getStream()->test(Twig_Token::BLOCK_END_TYPE)) {
            $properties = $this->parser->getExpressionParser()->parseExpression();
        }

        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);
        $body = $this->parser->subparse(function(Twig_Token $token) { return $token->test('end'.$this->getTag()); }, true);
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);

        $this->removeTextNodesRecursively($body);

        return new XlsDocumentNode($properties, $body, $token->getLine(), $this->getTag());
    }

    public function getTag()
    {
        return 'xlsdocument';
    }

    private function removeTextNodesRecursively(Twig_Node &$node) {
        foreach ($node->getIterator() as $key => $subNode)  {
            if ($subNode instanceof Twig_Node_Text) {
                $node->removeNode($key);
            }
            else if ($subNode instanceof Twig_Node && !($subNode instanceof XlsCellNode) && $subNode->count() > 0) {
                $this->removeTextNodesRecursively($subNode);
            }
        }
    }
}