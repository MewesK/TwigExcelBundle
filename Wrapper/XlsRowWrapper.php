<?php

namespace MewesK\TwigExcelBundle\Wrapper;

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
     * @var XlsSheetWrapper
     */
    protected $sheetWrapper;

    /**
     * @param array $context
     * @param XlsSheetWrapper $sheetWrapper
     */
    public function __construct(array $context, XlsSheetWrapper $sheetWrapper)
    {
        $this->context = $context;
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
            throw new \InvalidArgumentException(sprintf('Invalid index'));
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
