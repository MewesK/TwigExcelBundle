<?php

namespace MewesK\TwigExcelBundle\Tests\Twig\Mock;

/**
 * Class MockRequest
 * @package MewesK\TwigExcelBundle\Tests\Twig\Mock
 */
class MockRequest
{
    /**
     * @var string
     */
    private $requestFormat;

    /**
     * @param string $requestFormat
     */
    public function __construct($requestFormat)
    {
        $this->requestFormat = $requestFormat;
    }

    /**
     * @return string
     */
    public function getRequestFormat()
    {
        return $this->requestFormat;
    }
} 
