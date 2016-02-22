<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use MewesK\TwigExcelBundle\Twig\Node\XlsSheetNode;
use Twig_Node_Expression_Constant;
use Twig_Token;

/**
 * Class XlsSheetTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class XlsSheetTokenParser extends AbstractTokenParser
{
    /**
     * @param Twig_Token $token
     *
     * @return XlsSheetNode
     * @throws \Twig_Error_Syntax
     */
    public function parse(Twig_Token $token)
    {
        // parse attributes
        $title = new Twig_Node_Expression_Constant(null, $token->getLine());
        if (!$this->parser->getStream()->test(Twig_Token::PUNCTUATION_TYPE) && !$this->parser->getStream()->test(Twig_Token::BLOCK_END_TYPE)) {
            $title = $this->parser->getExpressionParser()->parseExpression();
        }
        $properties = $this->parseProperties($token);
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);

        // parse body
        $body = $this->parseBody();

        // return node
        return new XlsSheetNode($title, $properties, $body, $token->getLine(), $this->getTag());
    }

    /**
     * @return string
     */
    public function getTag()
    {
        return 'xlssheet';
    }
}
