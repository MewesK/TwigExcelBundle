<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use MewesK\TwigExcelBundle\Twig\Node\XlsLeftNode;
use Twig_Token;
use Twig_TokenParser;

/**
 * Class XlsLeftTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class XlsLeftTokenParser extends Twig_TokenParser
{
    /**
     * @param Twig_Token $token
     *
     * @return XlsLeftNode
     * @throws \Twig_Error_Syntax
     */
    public function parse(Twig_Token $token)
    {
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);
        $tokenParser = $this; // PHP 5.3 fix
        $body = $this->parser->subparse(function(Twig_Token $token) use ($tokenParser) { return $token->test('end'.$tokenParser->getTag()); }, true);
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);

        return new XlsLeftNode($body, $token->getLine(), $this->getTag());
    }

    /**
     * @return string
     */
    public function getTag()
    {
        return 'xlsleft';
    }
}
