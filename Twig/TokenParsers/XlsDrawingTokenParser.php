<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\TokenParsers;

use MewesK\PhpExcelTwigExtensionBundle\Twig\Nodes\XlsStyleNode;

class XlsDrawingTokenParser extends \Twig_TokenParser
{
    public function parse(\Twig_Token $token)
    {
        $path = $this->parser->getExpressionParser()->parseExpression();

        $properties = new \Twig_Node_Expression_Array([], $token->getLine());
        if (!$this->parser->getStream()->test(\Twig_Token::BLOCK_END_TYPE)) {
            $properties = $this->parser->getExpressionParser()->parseExpression();
        }

        $this->parser->getStream()->expect(\Twig_Token::BLOCK_END_TYPE);

        return new XlsStyleNode($path, $properties, $token->getLine(), $this->getTag());
    }

    public function getTag()
    {
        return 'xlsdrawing';
    }
}