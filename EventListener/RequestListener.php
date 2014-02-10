<?php

namespace MewesK\TwigExcelBundle\EventListener;

use Symfony\Component\HttpKernel\Event\GetResponseEvent;

class RequestListener
{
    public function onKernelRequest(GetResponseEvent $event)
    {
        $event->getRequest()->setFormat('csv', 'text/csv');
        $event->getRequest()->setFormat('xls', 'application/vnd.ms-excel');
        $event->getRequest()->setFormat('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $event->getRequest()->setFormat('pdf', ' application/pdf');
    }
}