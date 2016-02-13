<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use MewesK\TwigExcelBundle\Twig\NodeHelper;
use Twig_Error_Syntax;
use Twig_Node_Body;
use Twig_Node_Expression_Constant;
use Twig_Node_Macro;
use Twig_Token;
use Twig_TokenParser_Macro;

/**
 * Class XlsMacroTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class XlsMacroTokenParser extends Twig_TokenParser_Macro
{
    /**
     * Copy of the parent method to allow manipulating the Twig_Node_Macro instance.
     *
     * @param Twig_Token $token
     * @return \Twig_NodeInterface|void
     * @throws Twig_Error_Syntax
     */
    public function parse(Twig_Token $token)
    {
        $lineno = $token->getLine();
        $stream = $this->parser->getStream();
        $name = $stream->expect(Twig_Token::NAME_TYPE)->getValue();

        $arguments = $this->parser->getExpressionParser()->parseArguments(true, true);

        // fix macro context
        $arguments->setNode('phpExcel', new Twig_Node_Expression_Constant(null, null));

        $stream->expect(Twig_Token::BLOCK_END_TYPE);
        $this->parser->pushLocalScope();
        $body = $this->parser->subparse(array($this, 'decideBlockEnd'), true);
        if ($token = $stream->nextIf(Twig_Token::NAME_TYPE)) {
            $value = $token->getValue();

            if ($value != $name) {
                throw new Twig_Error_Syntax(sprintf('Expected endxlsmacro for macro "%s" (but "%s" given).', $name, $value), $stream->getCurrent()->getLine(), $stream->getFilename());
            }
        }
        $this->parser->popLocalScope();
        $stream->expect(Twig_Token::BLOCK_END_TYPE);

        // remove all unwanted text nodes
        NodeHelper::removeTextNodesRecursively($body, $this->parser);

        $macro = new Twig_Node_Macro($name, new Twig_Node_Body(array($body)), $arguments, $lineno, $this->getTag());

        // mark for syntax checks
        $macro->setAttribute('twigExcelBundle', true);

        $this->parser->setMacro($name, $macro);
    }

    /**
     * @param Twig_Token $token
     * @return bool
     */
    public function decideBlockEnd(Twig_Token $token)
    {
        return $token->test('endxlsmacro');
    }

    /**
     * @return string
     */
    public function getTag()
    {
        return 'xlsmacro';
    }
}
