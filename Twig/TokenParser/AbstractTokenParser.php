<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use Twig_Node_Expression_Array;
use Twig_Token;
use Twig_TokenParser;

/**
 * Class AbstractTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
abstract class AbstractTokenParser extends Twig_TokenParser
{
    /**
     * @param Twig_Token $token
     *
     * @return mixed|Twig_Node_Expression_Array|\Twig_Node_Expression_Conditional|\Twig_Node_Expression_GetAttr
     */
    protected function parseProperties(Twig_Token $token)
    {
        $properties = new Twig_Node_Expression_Array([], $token->getLine());
        if (!$this->parser->getStream()->test(Twig_Token::BLOCK_END_TYPE)) {
            $properties = $this->parser->getExpressionParser()->parseExpression();
        }

        return $properties;
    }

    /**
     * @param Twig_Token $token
     *
     * @return \Twig_Node
     * @throws \Twig_Error_Syntax
     */
    protected function parseBody(Twig_Token $token)
    {
        $tokenParser = $this; // PHP 5.3 fix
        $body = $this->parser->subparse(function(Twig_Token $token) use ($tokenParser) { return $token->test('end'.$tokenParser->getTag()); }, true);
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);

        return $body;
    }
}
