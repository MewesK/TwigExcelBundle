<?php

namespace MewesK\TwigExcelBundle\Tests;

/**
 * Class MockGlobalVariables
 *
 * @package MewesK\TwigExcelBundle\Tests
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
} 
