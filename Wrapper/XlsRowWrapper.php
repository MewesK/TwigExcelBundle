<?php

namespace MewesK\TwigExcelBundle\Wrapper;
use Twig_Environment;

/**
 * Class XlsSheetWrapper
 *
 * @package MewesK\TwigExcelBundle\Wrapper
 */
class XlsRowWrapper extends AbstractWrapper
{
    /**
     * @var array
     */
    protected $context;
    /**
     * @var Twig_Environment
     */
    protected $environment;
    /**
     * @var XlsSheetWrapper
     */
    protected $sheetWrapper;

    /**
     * XlsRowWrapper constructor.
     * @param array $context
     * @param Twig_Environment $environment
     * @param XlsSheetWrapper $sheetWrapper
     */
    public function __construct(array $context, Twig_Environment $environment, XlsSheetWrapper $sheetWrapper)
    {
        $this->context = $context;
        $this->environment = $environment;
        $this->sheetWrapper = $sheetWrapper;
    }

    /**
     * @param null|int $index
     */
    public function start($index = null)
    {
        if ($this->sheetWrapper->getObject() === null) {
            throw new \LogicException();
        }
        if ($index !== null && !is_int($index)) {
            throw new \InvalidArgumentException('Invalid index');
        }

        if ($index === null) {
            $this->sheetWrapper->increaseRow();
        } else {
            $this->sheetWrapper->setRow($index);
        }
    }

    public function end()
    {
        $this->sheetWrapper->setColumn(null);
    }
}
