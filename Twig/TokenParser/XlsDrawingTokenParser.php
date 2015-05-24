<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use MewesK\TwigExcelBundle\Twig\Node\XlsDrawingNode;
use Twig_Token;

/**
 * Class XlsDrawingTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class XlsDrawingTokenParser extends AbstractTokenParser
{
    /**
     * @param Twig_Token $token
     *
     * @return XlsDrawingNode
     * @throws \Twig_Error_Syntax
     */
    public function parse(Twig_Token $token)
    {
        // parse attributes
        $path = $this->parser->getExpressionParser()->parseExpression();
        $properties = $this->parseProperties($token);
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);

        // return node
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
