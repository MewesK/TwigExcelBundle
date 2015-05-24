<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use MewesK\TwigExcelBundle\Twig\Node\XlsFooterNode;
use Twig_Node_Expression_Constant;
use Twig_Token;

/**
 * Class XlsFooterTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class XlsFooterTokenParser extends AbstractTokenParser
{
    /**
     * @param Twig_Token $token
     *
     * @return XlsFooterNode
     * @throws \Twig_Error_Syntax
     */
    public function parse(Twig_Token $token)
    {
        // parse attributes
        $type = new Twig_Node_Expression_Constant('footer', $token->getLine());
        if (!$this->parser->getStream()->test(Twig_Token::PUNCTUATION_TYPE) && !$this->parser->getStream()->test(Twig_Token::BLOCK_END_TYPE)) {
            $type = $this->parser->getExpressionParser()->parseExpression();
        }
        $properties = $this->parseProperties($token);
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);

        // parse body
        $body = $this->parseBody();

        // return node
        return new XlsFooterNode($type, $properties, $body, $token->getLine(), $this->getTag());
    }

    /**
     * @return string
     */
    public function getTag()
    {
        return 'xlsfooter';
    }
}
