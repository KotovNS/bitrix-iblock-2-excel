<?php
require_once "ListExcel.php";

// $a = new ListExcel(["IBLOCK_ID"=>3, "ACTIVE"=>"Y", ">PROPERTY_PROP_5"=>100], ["NAME", "PROPERTY_PROP_6", "PROPERTY_PROP_5"], ["PROPERTY_PROP_5" => "ASC"], []);
// $a = new ListExcel(["IBLOCK_ID"=>3], ["NAME", "PROPERTY_PROP_6", "PROPERTY_PROP_5"], ["PROPERTY_PROP_6" => "ASC"], ["PROPERTY_PROP_6", "PROPERTY_PROP_5", "NAME"]);
$a = new ListExcel(["IBLOCK_ID"=>2], ["ID", "NAME"], ["ID" => "ASC"], ["ID", "NAME"]);

echo $a;
// echo "<hr>";
// sleep(1);

// $a->columnOrder = ["NAME", "PROPERTY_PROP_5", "PROPERTY_PROP_6"];
// $a->generateXLSXFromIBlock();

// echo $a;
// echo "<hr>";