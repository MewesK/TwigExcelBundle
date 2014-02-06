<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\TokenParser;

use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsFooterNode;
use Twig_Node_Expression_Array;
use Twig_Node_Expression_Constant;
use Twig_Token;
use Twig_TokenParser;

class XlsFooterTokenParser extends Twig_TokenParser
{
    public function parse(Twig_Token $token)
    {
        $type = new Twig_Node_Expression_Constant('header', $token->getLine());
        if (!$this->parser->getStream()->test(Twig_Token::BLOCK_END_TYPE)) {
            $type = $this->parser->getExpressionParser()->parseExpression();
        }
        $properties = new Twig_Node_Expression_Array(array(), $token->getLine());
        if (!$this->parser->getStream()->test(Twig_Token::BLOCK_END_TYPE)) {
            $properties = $this->parser->getExpressionParser()->parseExpression();
        }

        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);
        $tokenParser = $this; // PHP 5.3 fix
        $body = $this->parser->subparse(function(Twig_Token $token) use ($tokenParser) { return $token->test('end'.$tokenParser->getTag()); }, true);
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);

        return new XlsFooterNode($type, $properties, $body, $token->getLine(), $this->getTag());
    }

    public function getTag()
    {
        return 'xlsfooter';
    }
}