<?php

namespace MewesK\TwigExcelBundle\Wrapper;

/**
 * Class AbstractWrapper
 *
 * @package MewesK\TwigExcelBundle\Wrapper
 */
abstract class AbstractWrapper
{
    /**
     * @param array $properties
     * @param array $mappings
     */
    protected function setProperties(array $properties, array $mappings)
    {
        foreach ($properties as $key => $value) {
            if (array_key_exists($key, $mappings)) {
                if (is_array($value) && is_array($mappings) && $key !== 'style' && $key !== 'defaultStyle') {
                    if (array_key_exists('__multi', $mappings[$key]) && $mappings[$key]['__multi'] === true) {
                        foreach ($value as $_key => $_value) {
                            $this->setPropertiesByKey($_key, $_value, $mappings[$key]);
                        }
                    } else {
                        $this->setProperties($value, $mappings[$key]);
                    }
                } else {
                    $mappings[$key]($value);
                }
            }
        }
    }

    /**
     * @param string $key
     * @param array $properties
     * @param array $mappings
     */
    protected function setPropertiesByKey($key, array $properties, array $mappings)
    {
        foreach ($properties as $_key => $value) {
            if (array_key_exists($_key, $mappings)) {
                if (is_array($value)) {
                    $this->setPropertiesByKey($key, $value, $mappings[$_key]);
                } else {
                    $mappings[$_key]($key, $value);
                }
            }
        }
    }
}
