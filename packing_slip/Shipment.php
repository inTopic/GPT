<?php

/**
 * Copyright 3WebApps. All rights reserved.
 * https://3webapps.com
 */

declare(strict_types=1);

namespace Mage42\PackingSlip\Model\Order\Pdf;

use Magento\Customer\Api\GroupRepositoryInterface;
use Magento\Framework\Api\SearchCriteriaBuilder;
use Magento\Framework\Api\SortOrderBuilder;
use Magento\Framework\App\Config\ScopeConfigInterface;
use Magento\Framework\Stdlib\DateTime\TimezoneInterface;
use Magento\Sales\Api\Data\OrderInterface;
use Magento\Sales\Api\Data\OrderItemInterface;
use Magento\Sales\Api\OrderItemRepositoryInterface;
use Magento\Sales\Model\Order;
use Magento\Sales\Model\Order\Pdf\Config;
use Magento\Sales\Model\Order\Pdf\Total\Factory;
use Magento\Sales\Model\Order\Pdf\ItemsFactory;
use Magento\Store\Model\App\Emulation;
use Magento\Sales\Model\Order\Pdf\Shipment as SalesShipment;
use Magento\Store\Model\ScopeInterface;
use Magento\Store\Model\Store;
use Magento\Payment\Helper\Data;
use Magento\Framework\Stdlib\StringUtils;
use Magento\Framework\Filesystem;
use Magento\Framework\Translate\Inline\StateInterface;
use Magento\Sales\Model\Order\Address\Renderer;
use Magento\Store\Model\StoreManagerInterface;
use Zend_Pdf;
use Zend_Pdf_Exception;
use Zend_Pdf_Page;

class Shipment extends SalesShipment
{
    /**
     * @var Emulation
     */
    private Emulation  $appEmulation;
    private SearchCriteriaBuilder $searchCriteriaBuilder;
    private OrderItemRepositoryInterface $orderItemRepository;
    private SortOrderBuilder $sortOrderBuilder;
    private TimezoneInterface $timezone;
    private GroupRepositoryInterface $groupRepository;

    public function __construct(
        Data $paymentData,
        StringUtils $string,
        ScopeConfigInterface $scopeConfig,
        Filesystem $filesystem,
        Config $pdfConfig,
        Factory $pdfTotalFactory,
        ItemsFactory $pdfItemsFactory,
        TimezoneInterface $localeDate,
        StateInterface $inlineTranslation,
        Renderer $addressRenderer,
        StoreManagerInterface $storeManager,
        Emulation $appEmulation,
        SearchCriteriaBuilder $searchCriteriaBuilder,
        OrderItemRepositoryInterface $orderItemRepository,
        SortOrderBuilder $sortOrderBuilder,
        GroupRepositoryInterface $groupRepository,
        array $data = []
    ) {
        parent::__construct($paymentData, $string, $scopeConfig, $filesystem,
            $pdfConfig, $pdfTotalFactory, $pdfItemsFactory, $localeDate,
            $inlineTranslation, $addressRenderer, $storeManager, $appEmulation,
            $data);
        $this->groupRepository = $groupRepository;
        $this->appEmulation = $appEmulation;
        $this->searchCriteriaBuilder = $searchCriteriaBuilder;
        $this->orderItemRepository = $orderItemRepository;
        $this->sortOrderBuilder = $sortOrderBuilder;
    }

    /**
     * Draw table header for product items
     *
     * @param Zend_Pdf_Page $page
     *
     * @return void
     * @throws Zend_Pdf_Exception
     */
    protected function _drawHeader(Zend_Pdf_Page $page): void
    {
        $this->_setFontRegular($page, 10);

        $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0.5));
        $page->setLineWidth(0);
        $top = 700;
        $page->drawRectangle(25, $top, 570, $top - 15);

        $page->setFillColor(new \Zend_Pdf_Color_GrayScale(1));
        $page->drawText("Beschrijving", 30, 690);

        $page->setFillColor(new \Zend_Pdf_Color_GrayScale(1));
        $page->drawText("Gewicht", 205, 690);

        $page->setFillColor(new \Zend_Pdf_Color_GrayScale(1));
        $page->drawText("Artikelcode", 320, 690);

        $page->setFillColor(new \Zend_Pdf_Color_GrayScale(1));
        $page->drawText("SKU", 420, 690);

        $page->setFillColor(new \Zend_Pdf_Color_GrayScale(1));
        $page->drawText("Aantal", 490, 690);
    }

    /**
     * Return PDF document
     *
     * @param \Magento\Sales\Model\Order\Shipment[] $shipments
     *
     * @return Zend_Pdf
     * @throws Zend_Pdf_Exception
     */
    public function getPdf($shipments = []): Zend_Pdf
    {
        $this->_beforeGetPdf();
        $this->_initRenderer('shipment');

        $pdf = new Zend_Pdf();
        $this->_setPdf($pdf);
        $style = new \Zend_Pdf_Style();
        $this->_setFontBold($style, 10);
        foreach ($shipments as $shipment) {
            if ($shipment->getStoreId()) {
                $this->appEmulation->startEnvironmentEmulation(
                    $shipment->getStoreId(),
                    \Magento\Framework\App\Area::AREA_FRONTEND,
                    true
                );
                $this->_storeManager->setCurrentStore($shipment->getStoreId());
            }
            $page = $this->newPage();
            $order = $shipment->getOrder();

            /* Add image */
            $this->drawLogo($page, $shipment->getStore());

            /** Add order number */
            $page->setFont(\Zend_Pdf_Font::fontWithName(\Zend_Pdf_Font::FONT_HELVETICA_BOLD), 16);
            $page->drawText("Pakbon", 25, 780);
            $this->_setFontRegular($page, 12);
            $page->drawText("Order: ".$order->getIncrementId(), 25, 765);

            $this->_setFontBold($page, 10);

            /* Delivery method */
            $this->_setFontBold($page, 8);
            $page->drawText("Verzendmethode", 25, 725);
            $this->_setFontRegular($page, 8);

            $shippingDescription = $order->getShippingDescription();
            $shippingMethod = $shippingDescription;
            if (str_contains($shippingDescription, ' - ')) {
                $shippingMethod =
                    explode(" - ", $shippingDescription)[1];
            }
            $page->drawText($shippingMethod, 25, 715);

            /* Payment method */
            $this->_setFontBold($page, 8);
            $page->drawText("Betaalmethode", 205, 725);
            $this->_setFontRegular($page, 8);
            $page->drawText(preg_replace("/\([^)]+\)/","",$order->getPayment()->getMethodInstance()->getTitle() ?? ''), 205, 715);

            /* Order date */
            $this->_setFontBold($page, 8);
            $page->drawText("Besteldatum", 320, 725);
            $this->_setFontRegular($page, 8);

            $dateText = $this->_localeDate
                ->date(strtotime($order->getCreatedAt()))
                ->format('l d M Y');

            $page->drawText($dateText, 320, 715);

            /* Add table */
            $this->_drawHeader($page);
            /* Add body */
            $this->drawItems($page, $order);

            $height = 670 - count($order->getItems()) * 20;

            /** Addresses */
            $billing = $order->getBillingAddress();
            $shipping = $order->getShippingAddress();

            $customerGroup = '';
            if ($order->getCustomerGroupId()) {
            $customerGroup = $this->groupRepository
                ->getById($order->getCustomerGroupId())
                ->getCode();
            }

            $this->_setFontBold($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText('Factuuradres:', 25, $height-10);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText($customerGroup, 25, $height-22);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText(ucfirst($billing->getFirstname()).' '.ucfirst($billing->getLastname()), 25, $height-34);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText(implode(' ', $billing->getStreet()), 25, $height-46);

            $postCode = $billing->getPostcode() ?? "";
            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText($postCode, 25, $height-58);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText($billing->getCity(), 25, $height-70);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText($billing->getCountryId(), 25, $height-82);

            $this->_setFontBold($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText('Contact informatie:', 205, $height-10);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText($order->getCustomerEmail(), 205, $height-22);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText($billing->getTelephone(), 205, $height-34);

            $this->_setFontBold($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText('Afleveradres:', 320, $height-10);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText(ucfirst($shipping->getFirstname()).' '.ucfirst($shipping->getLastname()), 320, $height-22);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText(implode(' ', $shipping->getStreet()), 320, $height-34);

            $postCode = $shipping->getPostcode() ?? "";
            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText($postCode, 320, $height-46);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText($shipping->getCity(), 320, $height-58);

            $this->_setFontRegular($page, 8);
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            $page->drawText($shipping->getCountryId(), 320, $height-70);



            $this->insertBoldOrderComment($page, $order, $height);


            end($pdf->pages);


            if ($shipment->getStoreId()) {
                $this->appEmulation->stopEnvironmentEmulation();
            }
        }
        $this->_afterGetPdf();
        return $pdf;
    }
    /**
     * Insert bold order comment into the PDF
     */
    protected function insertBoldOrderComment($page, $order, $height)
    {
        $height = $height - 100;
        $boldOrderComment = $order->getData('bold_order_comment');

        $diff = 12;
        if ($boldOrderComment) {
            $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
            while(strlen($boldOrderComment) > 0) {
                if (strlen($boldOrderComment) > 100) {
                    $end = strrpos(substr($boldOrderComment, 0, 100), ' ');
                    $end = $end ? $end : 100;
                    $line = substr($boldOrderComment, 0, $end);
                    $boldOrderComment = substr($boldOrderComment, $end+1);
                } else {
                    $line = $boldOrderComment;
                    $boldOrderComment = '';
                }
                $page->drawText($line, 25, $height-$diff, 'UTF-8');
                $diff = $diff+12;
            }
        }
    }

    /**
     * Draw Item process
     *
     * @param Zend_Pdf_Page                         $page
     * @param Order                                 $order
     *
     * @return Zend_Pdf_Page
     * @throws Zend_Pdf_Exception
     */
    protected function drawItems(
        Zend_Pdf_Page $page,
        \Magento\Sales\Model\Order $order
    ): Zend_Pdf_Page {
        $height = 670;

        $orderItems = $this->getItemsSorted($order);
        $fillColor = [0.7, 0.9];

        $counter = 0;
        foreach ($orderItems as $item) {
            $type = $item->getProductType();

            $indent = 30;
            if($item->getParentItemId() !== null)
                $indent = 40;


            if(in_array($type, ['configurable', 'bundle'])) {

                $page->setFillColor(new \Zend_Pdf_Color_GrayScale($fillColor[$counter%2]));
                $page->setLineWidth(0);
                $page->drawRectangle(25, $height + 15, 570, $height - 5);

                $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
                $page->drawText($item->getName(), $indent, $height+3);

                $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
                $page->drawText(substr($item->getSku(), 0, 4), 320, $height+3);

                $height = $height-20;

            } else {
                $page->setFillColor(new \Zend_Pdf_Color_GrayScale($fillColor[$counter%2]));
                $page->setLineWidth(0);
                $page->drawRectangle(25, $height + 15, 570, $height - 5);

                $page->setFillColor(new \Zend_Pdf_Color_GrayScale(1));
                $page->setLineWidth(0);
                $page->drawRectangle(550, $height + 10, 560, $height);

                if ($item->getProduct()) {
                    $attributeText = $item->getProduct()->getAttributeText('bs_weight');
                    if(isset($attributeText)) {
                        $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
                        $page->drawText($attributeText, 205, $height+3);
                    }
                }

                $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
                $page->drawText(substr($item->getName(),0,35), $indent, $height+3);

                $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
                $page->drawText(substr($item->getSku(), 0, 4), 320, $height+3);

                $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
                $page->drawText($item->getSku(), 420, $height+3);

                $page->setFillColor(new \Zend_Pdf_Color_GrayScale(0));
                $page->drawText(intval($item->getData('qty_ordered')).'x', 490, $height+3);

                $height = $height-20;
            }
            $counter++;
        }

        return $page;
    }

    /**
     * Retrive order items sorted by SKU asc.
     *
     * @param OrderInterface $order
     *
     * @return OrderItemInterface[]
     */
    private function getItemsSorted(OrderInterface $order): array
    {
        $sortOrder = $this->sortOrderBuilder
            ->setField('sku')
            ->setAscendingDirection()
            ->create();

        $searchCriteria = $this->searchCriteriaBuilder
            ->addFilter('order_id', $order->getEntityId())
            ->addSortOrder($sortOrder)
            ->create();

        return $this->orderItemRepository
            ->getList($searchCriteria)
            ->getItems();
    }

    private function drawLogo(Zend_Pdf_Page $page, Store $store)
    {
        $image = $this->_scopeConfig->getValue(
            'sales/identity/logo',
            ScopeInterface::SCOPE_STORE,
            $store
        );
        $imagePath = '/sales/store/logo/' . $image;
        $imagePath = $this->_mediaDirectory->getAbsolutePath($imagePath);

        $logo = \Zend_Pdf_Image::imageWithPath($imagePath);

        $h = 425;
        $b = 700;
        $page->drawImage($logo, $h, $b, $h + 157.8, $b + 139.25);
    }

    /**
     * Set font as regular
     *
     * @param  \Zend_Pdf_Page $object
     * @param  int $size
     * @return \Zend_Pdf_Resource_Font
     */
    protected function _setFontRegular($object, $size = 7)
    {
        $font = \Zend_Pdf_Font::fontWithName(
            \Zend_Pdf_Font::FONT_HELVETICA
        );
        $object->setFont($font, $size);
        return $font;
    }

    /**
     * Set font as bold
     *
     * @param  \Zend_Pdf_Page $object
     * @param  int $size
     * @return \Zend_Pdf_Resource_Font
     */
    protected function _setFontBold($object, $size = 7)
    {
        $font = \Zend_Pdf_Font::fontWithName(
            \Zend_Pdf_Font::FONT_HELVETICA_BOLD
        );
        $object->setFont($font, $size);
        return $font;
    }

}
