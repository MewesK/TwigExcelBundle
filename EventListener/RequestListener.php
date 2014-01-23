<?php

namespace MewesK\PhpExcelTwigExtensionBundle\EventListener;

use Symfony\Component\HttpKernel\Event\GetResponseEvent;

class RequestListener
{
    public function onKernelRequest(GetResponseEvent $event)
    {
        $event->getRequest()->setFormat('xls', 'application/vnd.ms-excel');
    }
}