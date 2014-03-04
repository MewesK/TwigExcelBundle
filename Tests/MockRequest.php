<?php

namespace MewesK\TwigExcelBundle\Tests;

class MockRequest {
    /**
     * @var string
     */
    private $requestFormat;

    public function __construct($requestFormat) {
        $this->requestFormat = $requestFormat;
    }

    public function getRequestFormat() {
        return $this->requestFormat;
    }
} 