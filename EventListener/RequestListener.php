<?php

namespace MewesK\PhpExcelTwigExtensionBundle\EventListener;

use Symfony\Component\HttpKernel\Event\GetResponseEvent;

class RequestListener
{
    public function onKernelRequest(GetResponseEvent $event)
    {
        $event->getRequest()->setFormat('csv', 'text/csv'); // CSV
        $event->getRequest()->setFormat('xls', 'application/vnd.ms-excel'); // Excel5
        $event->getRequest()->setFormat('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); // Excel2007
        $event->getRequest()->setFormat('pdf', ' application/pdf'); // PDF
    }
}