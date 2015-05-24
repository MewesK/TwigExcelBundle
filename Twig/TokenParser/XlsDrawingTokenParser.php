<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use MewesK\TwigExcelBundle\Twig\Node\XlsDrawingNode;
use Twig_Node_Expression_Array;
use Twig_Token;
use Twig_TokenParser;

/**
 * Class XlsDrawingTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class XlsDrawingTokenParser extends Twig_TokenParser
{
    /**
     * @param Twig_Token $token
     *
     * @return XlsDrawingNode
     * @throws \Twig_Error_Syntax
     */
    public function parse(Twig_Token $token)
    {
        $path = $this->parser->getExpressionParser()->parseExpression();

        $properties = new Twig_Node_Expression_Array([], $token->getLine());
        if (!$this->parser->getStream()->test(Twig_Token::BLOCK_END_TYPE)) {
            $properties = $this->parser->getExpressionParser()->parseExpression();
        }

        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);

        return new XlsDrawingNode($path, $properties, $token->getLine(), $this->getTag());
    }

    /**
     * @return string
     */
    public function getTag()
    {
        return 'xlsdrawing';
    }
}
