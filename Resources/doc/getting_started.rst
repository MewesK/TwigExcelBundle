Getting started
===============

Step 1: Create your controller
------------------------------

.. code-block:: php

    <?php
    // src/Acme/HelloBundle/Controller/HelloController.php

    namespace Acme\HelloBundle\Controller;

    use Sensio\Bundle\FrameworkExtraBundle\Configuration\Route;
    use Sensio\Bundle\FrameworkExtraBundle\Configuration\Template;
    use Symfony\Component\HttpFoundation\Response;

    class HelloController
    {
        /**
         * @Route("/hello.{_format}", defaults={"_format"="xls"}, requirements={"_format"="csv|xls|xlsx"})
         * @Template("AcmeHelloBundle:Hello:index.xls.twig")
         */
        public function indexAction($name)
        {
            return ['data' => ['La', 'Le', 'Lu']];
        }
    }

Step 2: Create your template
----------------------------

.. code-block:: twig

    {# src/Acme/HelloBundle/Resources/views/Hello/index.xls.twig #}

    {% xlsdocument %}
        {% xlssheet 'Worksheet' %}
            {% xlsrow %}
                {% xlscell { style: { font: { size: '18' } } } %}Values{% endxlscell %}
            {% endxlsrow %}
            {% for value in data %}
                {% xlsrow %}
                    {% xlscell %}{{ value }}{% endxlscell %}
                {% endxlsrow %}
            {% endfor %}
        {% endxlssheet %}
    {% endxlsdocument %}
