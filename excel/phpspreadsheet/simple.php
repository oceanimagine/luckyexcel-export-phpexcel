<?php

require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Hello World !');

$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
$drawing->setWorksheet($spreadsheet->getActiveSheet());
$drawing->setPath(__DIR__ . '/../images/28AF5666BC67160220220713195709.jpeg');
$drawing->setWidthAndHeight(158, 72);
$drawing->setResizeProportional(true);

$drawing->setOffsetX(400);    // this is how
$drawing->setOffsetY(300);    // this is how
$drawing->setOffsetX2(400);    // this is how
$drawing->setOffsetY2(300);    // this is how

$writer = new Xlsx($spreadsheet);
$writer->save('hello world.xlsx');