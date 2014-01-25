<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\TokenParsers;

use MewesK\PhpExcelTwigExtensionBundle\Twig\Nodes\XlsRowNode;

class XlsRowTokenParser extends \Twig_TokenParser
{
    public function parse(\Twig_Token $token)
    {
        $index = $this->parser->getExpressionParser()->parseExpression();

        $properties = new \Twig_Node_Expression_Array([], $token->getLine());
        if (!$this->parser->getStream()->test(\Twig_Token::BLOCK_END_TYPE)) {
            $properties = $this->parser->getExpressionParser()->parseExpression();
        }

        $this->parser->getStream()->expect(\Twig_Token::BLOCK_END_TYPE);
        $body = $this->parser->subparse(function(\Twig_Token $token) { return $token->test('endxlsrow'); }, true);
        $this->parser->getStream()->expect(\Twig_Token::BLOCK_END_TYPE);

        $this->checkSyntaxErrorsRecursively($body);

        return new XlsRowNode($index, $properties, $body, $token->getLine(), $this->getTag());
    }

    public function getTag()
    {
        return 'xlsrow';
    }

    private function checkSyntaxErrorsRecursively(\Twig_Node $node) {
        foreach ($node->getIterator() as $subNode) {
            if ($subNode instanceof XlsRowNode) {
                throw new \LogicException(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($subNode), get_class($node))
                );
            }
            if ($subNode instanceof \Twig_Node && $subNode->count() > 0) {
                $this->checkSyntaxErrorsRecursively($subNode);
            }
        }
    }
}