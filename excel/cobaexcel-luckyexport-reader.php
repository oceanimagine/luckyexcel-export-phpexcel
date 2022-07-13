<?php

require_once __DIR__.'/phpexcel/PHPExcel.php';

$inputFileName = __DIR__ . '/phpexcel-output/SAMPLEBORDERWITHCELLNAME.xlsx';

//  Read your Excel workbook
try {
    $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load($inputFileName);
} catch(Exception $e) {
    die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
}

$i = 0;
echo $objPHPExcel->getSheetCount() . "\n";

$name_array = (array) $objPHPExcel->getNamedRanges();
$name_keys = array_keys($name_array);

print_r($name_keys);
for($i = 0; $i < $objPHPExcel->getSheetCount(); $i++){
    $objPHPExcel->setActiveSheetIndex($i);
    $objWorksheet = $objPHPExcel->getActiveSheet();
}

for($j = 0; $j < sizeof($name_keys); $j++){
    echo $objPHPExcel->getNamedRange($name_keys[$j])->getWorksheet()->getTitle() . " - " . $name_keys[$j] . " - " . $objPHPExcel->getNamedRange($name_keys[$j])->getRange();
    echo "\n";
}