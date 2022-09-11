<?php
require($_SERVER["DOCUMENT_ROOT"]."/bitrix/header.php");
require_once($_SERVER["DOCUMENT_ROOT"]."/Classes/PHPExcel.php");

/**
 * Класс для экспорта элементов инфоблока в эксель. Создаёт объект, содержащий вытянутые данные, фильтры и путь к сгенерированному файлу
 * $docPath - путь к документу
 * 
 * Аргументы конструктора:
 * $arFilter - (массив) - фильтр
 * $arSelect - (массив) - выводимые свойства
 * $arOrder - (массив) - сортировка
 * $columnOrder - (массив) - порядок записи свойств в таблицу excel
 * 
 * Формулировка названий свойств в $arSelect и $columnOrder должны быть одинаковыми
 */
class ListExcel {
    public $arFilter, $arSelect, $arOrder, $columnOrder, $iBlockData, $sortedData, $docPath;

    public function __construct($arFilter = [], $arSelect = [], $arOrder = [], $columnOrder = []) {
        $this->arFilter = $arFilter;
        $this->arSelect = $arSelect;
        $this->arOrder = $arOrder;
        $this->columnOrder = $columnOrder;

        $this->generateXLSXFromIBlock();
    }

    /**
     * С помощью этого метода можно перегенерировать документ вручную на случай изменения каких-нибудь свойств объекта
     */
    public function generateXLSXFromIBlock() {
        $this->iBlockData = $this->getList();
        $this->sortedData = $this->sortData();
        $this->docPath = $this->createXLSX();
    }

    /**
     * Возвращает неупорядоченный массив элементов по введённым фильтрам
     */
    private function getList() {
        // Для корректной работы битриксового GetList добавляем id инфоблока и элемента
        $arSelect = $this->arSelect;
        if (!in_array("IBLOCK_ID", $arSelect)) {
            $arSelect[] = "IBLOCK_ID";
        }
        if (!in_array("ID", $arSelect)) {
            $arSelect[] = "ID";
        }

        $data = [];

        $rows = CIBlockElement::GetList(
            $this->arOrder,
            $this->arFilter,
            false,
            [],
            $arSelect
        );

        while ($row = $rows->Fetch()) {
            $data[] = $row;
        }

        return $data;
    }

    /**
     * Возвращает упорядоченный массив элементов для записи в файл
     */
    private function sortData() {
        $data = [];
        $index = 0;
        if ($this->columnOrder == []) {
            $this->columnOrder = $this->arSelect;
        }
        foreach ($this->columnOrder as $col) {
            foreach ($this->iBlockData as $row) {
                if (strpos($col, "ROPERTY_")) {
                    $data[$index][$col] = $row[$col."_VALUE"];
                }
                else {
                    $data[$index][$col] = $row[$col];
                }
                $index++;
            }
            $index = 0;
        }

        return $data;
    }

    /**
     * Генерирует .xlsx-файл с полученными данными.
     * Возвращает путь к этому файлу
     */
    private function createXLSX() {
        $xlsx = new PHPExcel();
        $xlsx->setActiveSheetIndex(0);
        $sheet = $xlsx->getActiveSheet();
        $sheet->setTitle("Элементы инфоблока $this->iBlockID");

        if ($this->columnOrder) {
            $row = 1;
            $col = 0;
            foreach ($this->columnOrder as $label) {
                $sheet->setCellValueByColumnAndRow($col, $row, $label);
                $col++;
            }
        }
        else {
            $row = 1;
            $col = 0;
            foreach ($this->sortedData[0] as $label=>$value) {
                $sheet->setCellValueByColumnAndRow($col, $row, $label);
                $col++;
            }
        }

        $row = 2;
        $col = 0;
        foreach ($this->sortedData as $rows) {
            foreach ($rows as $value) {
                $sheet->setCellValueByColumnAndRow($col, $row, $value);
                $sheet->getColumnDimensionByColumn($col)->setAutoSize(true);
                $col++;
            }
            $row++;
            $col = 0;

        }

        $docPath .= __DIR__."/files/GetList_".date("Ymd_His").".xlsx";
        $objWriter = new PHPExcel_Writer_Excel2007($xlsx);
        $objWriter->save($docPath);

        return $docPath;
    }

    public function __toString() {
        return $this->docPath;
    }
}