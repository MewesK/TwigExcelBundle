<?php
namespace MewesK\TwigExcelBundle\Tests;

class MockGlobalVariables {
    /**
     * @var string
     */
    private $requestFormat;

    public function __construct($requestFormat) {
        $this->requestFormat = $requestFormat;
    }

    public function getRequest() {
        return new MockRequest($this->requestFormat);
    }
} 