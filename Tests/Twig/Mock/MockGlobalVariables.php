<?php

namespace MewesK\TwigExcelBundle\Tests\Twig\Mock;

/**
 * Class MockGlobalVariables
 * @package MewesK\TwigExcelBundle\Tests\Twig\Mock
 */
class MockGlobalVariables
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
     * @return MockRequest
     */
    public function getRequest()
    {
        return new MockRequest($this->requestFormat);
    }

    /**
     * @return string
     */
    public function getRequestFormat()
    {
        return $this->requestFormat;
    }
} 
