<?php
ob_start();
require_once __DIR__.'/phpexcel/PHPExcel.php';
if(!isset($_SERVER)){
    echo "Something Wrong With Your PHP Instalation.";
    exit();
}

function do_increment($param, $type = "HEX"){
    if($type == "ALP"){
        $huruf_angka = array(
            "A" =>  0,
            "B" =>  1,
            "C" =>  2,
            "D" =>  3,
            "E" =>  4,
            "F" =>  5,
            "G" =>  6,
            "H" =>  7,
            "I" =>  8,
            "J" =>  9,
            "K" => 10,
            "L" => 11,
            "M" => 12,
            "N" => 13,
            "O" => 14,
            "P" => 15,
            "Q" => 16,
            "R" => 17,
            "S" => 18,
            "T" => 19,
            "U" => 20,
            "V" => 21,
            "W" => 22,
            "X" => 23,
            "Y" => 24,
            "Z" => 25
        );
        $angka_huruf = array(
            0  => "A",
            1  => "B",
            2  => "C",
            3  => "D",
            4  => "E",
            5  => "F",
            6  => "G",
            7  => "H",
            8  => "I",
            9  => "J",
            10 => "K",
            11 => "L",
            12 => "M",
            13 => "N",
            14 => "O",
            15 => "P",
            16 => "Q",
            17 => "R",
            18 => "S",
            19 => "T",
            20 => "U",
            21 => "V",
            22 => "W",
            23 => "X",
            24 => "Y",
            25 => "Z",
        );
        $increment = 0;
    }
    if($type == "HEX"){
        $huruf_angka = array(
            "0" =>  0,
            "1" =>  1,
            "2" =>  2,
            "3" =>  3,
            "4" =>  4,
            "5" =>  5,
            "6" =>  6,
            "7" =>  7,
            "8" =>  8,
            "9" =>  9,
            "A" => 10,
            "B" => 11,
            "C" => 12,
            "D" => 13,
            "E" => 14,
            "F" => 15
        );
        $angka_huruf = array(
            0  => "0",
            1  => "1",
            2  => "2",
            3  => "3",
            4  => "4",
            5  => "5",
            6  => "6",
            7  => "7",
            8  => "8",
            9  => "9",
            10 => "A",
            11 => "B",
            12 => "C",
            13 => "D",
            14 => "E",
            15 => "F"
        );
        $increment = 1;
    }
    if($type == "DEC"){
        $huruf_angka = array(
            "0" =>  0,
            "1" =>  1,
            "2" =>  2,
            "3" =>  3,
            "4" =>  4,
            "5" =>  5,
            "6" =>  6,
            "7" =>  7,
            "8" =>  8,
            "9" =>  9
        );
        $angka_huruf = array(
            0  => "0",
            1  => "1",
            2  => "2",
            3  => "3",
            4  => "4",
            5  => "5",
            6  => "6",
            7  => "7",
            8  => "8",
            9  => "9"
        );
        $increment = 1;
    }
    if($type == "OCT"){
        $huruf_angka = array(
            "0" =>  0,
            "1" =>  1,
            "2" =>  2,
            "3" =>  3,
            "4" =>  4,
            "5" =>  5,
            "6" =>  6,
            "7" =>  7
        );
        $angka_huruf = array(
            0  => "0",
            1  => "1",
            2  => "2",
            3  => "3",
            4  => "4",
            5  => "5",
            6  => "6",
            7  => "7"
        );
        $increment = 1;
    }
    if($type == "BIN"){
        $huruf_angka = array(
            "0" =>  0,
            "1" =>  1
        );
        $angka_huruf = array(
            0  => "0",
            1  => "1"
        );
        $increment = 1;
    }
    $param_string = (string) $param;
    $result = "";
    $result_reserve = "";
    $after_f = 0;
    $count_f = 0;
    $keys = array_keys($huruf_angka);
    $last = $keys[sizeof($keys) - 1];
    $begin = $keys[0];
    for($i = strlen($param_string) - 1; $i >= 0; $i--){
        if(!$after_f){
            if($param_string{$i} == $last){
                $result_reserve = $result_reserve . $begin;
                $count_f++;
            } else {
                $angka = $huruf_angka[$param_string{$i}];
                $huruf = $angka_huruf[$angka + 1];
                $result_reserve = $result_reserve . $huruf;
                $after_f = 1;
            }
        } else {
            $result_reserve = $result_reserve . $param_string{$i};
        }
    }
    if($count_f == strlen($param_string)){
        $result = $keys[0 + $increment] . $result_reserve;
    } else {
        for($i = strlen($result_reserve) - 1; $i >= 0; $i--){
            $result = $result . $result_reserve{$i};
        }
    }
    return $result;
}

function convert_alphabet($param){
    $param_hex = 0;
    $param_dec = 0;
    $param_oct = 0;
    $param_bin = 0;
    $param_alp = "A";
    for($i = 0; $i < (int) $param; $i++){
        $param_alp = do_increment($param_alp, "ALP");
    }
    return array(
        $param_alp,
        $param_bin,
        $param_oct,
        $param_dec,
        $param_hex
    );
}

if(isset($_SERVER) && isset($_SERVER['REQUEST_METHOD']) && $_SERVER['REQUEST_METHOD'] == "POST"){
    echo "MASUK POST\n";
    $json_all = json_decode(json_encode($_POST['json_luckyexcel']));
    // file_put_contents("JSON_RAW", print_r($json_all, true));
    // file_put_contents("POST_RAW", json_encode($_POST['json_luckyexcel']));
    echo $json_all[0]->name . "\n";
    echo convert_alphabet((string)$json_all[0]->celldata[0]->c)[0] . "\n";
    echo $json_all[0]->celldata[0]->r + 1 . "\n";
    // exit();
    require_once __DIR__.'/phpexcel/PHPExcel.php';
    
    $style_border_top = array(
        'borders' => array(
            'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN)
        )
    );
    $style_border_bottom = array(
        'borders' => array(
            'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN)
        )
    );
    $style_border_right = array(
        'borders' => array(
            'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN)
        )
    );
    $style_border_left = array(
        'borders' => array(
            'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN)
        )
    );
    $style_align_center = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
        ) 
    ); 
    $style_align_right = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
        ) 
    );
    $style_vertical_middle = array(
        'alignment' => array(
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
        ) 
    );
    
    $excel = new PHPExcel();
    $excel->getProperties()->setCreator('Luckyexcel YKKBI')
            ->setLastModifiedBy('Luckyexcel YKKBI')
            ->setTitle("Luckyexcel YKKBI")
            ->setSubject("Luckyexcel YKKBI")
            ->setDescription("Luckyexcel YKKBI")
            ->setKeywords("Luckyexcel YKKBI");
    $sheet = 0;
    for($i = 0; $i < sizeof($json_all); $i++){
        if($i > 0){
            $excel->createSheet($i);
        }
        $newsheet = $excel->setActiveSheetIndex($i);
        $newsheet->setTitle($json_all[$i]->name);
        for($j = 0; isset($json_all[$i]->celldata) && $j < sizeof($json_all[$i]->celldata); $j++){
            if(isset($json_all[$i]->celldata[$j]->v->f)){
                $newsheet->setCellValue(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1), str_replace(" ", "+", $json_all[$i]->celldata[$j]->v->f));
                $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray($style_vertical_middle);
                if(isset($json_all[$i]->celldata[$j]->v->ht)){
                    if($json_all[$i]->celldata[$j]->v->ht == 0){
                        $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray($style_align_center);
                    }
                    if($json_all[$i]->celldata[$j]->v->ht == 2){
                        $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray($style_align_right);
                    }
                }
                if(isset($json_all[$i]->celldata[$j]->v->bl) && $json_all[$i]->celldata[$j]->v->bl){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->setBold(true);
                }
                if(isset($json_all[$i]->celldata[$j]->v->it) && $json_all[$i]->celldata[$j]->v->it){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->setItalic(true);
                }
                if(isset($json_all[$i]->celldata[$j]->v->fc)){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->getColor()->setRGB(substr($json_all[$i]->celldata[$j]->v->fc, 1));
                }
                if(isset($json_all[$i]->celldata[$j]->v->fs)){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->setSize($json_all[$i]->celldata[$j]->v->fs);
                }
                if(isset($json_all[$i]->celldata[$j]->v->ff)){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray(array("font" => array("name" => $json_all[$i]->celldata[$j]->v->ff)));
                }
            } 
            else if(isset($json_all[$i]->celldata[$j]->v->v)){
                $newsheet->setCellValue(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1), $json_all[$i]->celldata[$j]->v->v);
                $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray($style_vertical_middle);
                if(isset($json_all[$i]->celldata[$j]->v->ht)){
                    if($json_all[$i]->celldata[$j]->v->ht == 0){
                        $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray($style_align_center);
                    }
                    if($json_all[$i]->celldata[$j]->v->ht == 2){
                        $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray($style_align_right);
                    }
                }
                if(isset($json_all[$i]->celldata[$j]->v->bl) && $json_all[$i]->celldata[$j]->v->bl){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->setBold(true);
                }
                if(isset($json_all[$i]->celldata[$j]->v->it) && $json_all[$i]->celldata[$j]->v->it){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->setItalic(true);
                }
                if(isset($json_all[$i]->celldata[$j]->v->fc)){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->getColor()->setRGB(substr($json_all[$i]->celldata[$j]->v->fc, 1));
                }
                if(isset($json_all[$i]->celldata[$j]->v->fs)){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->setSize($json_all[$i]->celldata[$j]->v->fs);
                }
                if(isset($json_all[$i]->celldata[$j]->v->ff)){
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray(array("font" => array("name" => $json_all[$i]->celldata[$j]->v->ff)));
                }
            }
            
            if(isset($json_all[$i]->celldata[$j]->v->ct)){
                if(isset($json_all[$i]->celldata[$j]->v->bg)){
                    // echo "BG\n";
                    $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray(array(
                        "fill" => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => substr($json_all[$i]->celldata[$j]->v->bg, 1))
                        )
                    ));
                }
                if(isset($json_all[$i]->celldata[$j]->v->ct->s)){
                    $value_all = "";
                    for($h = 0; $h < sizeof($json_all[$i]->celldata[$j]->v->ct->s); $h++){
                        $font_info = $json_all[$i]->celldata[$j]->v->ct->s[$h];
                        if(isset($font_info->v)){
                            $value_all = $value_all . $font_info->v;
                            $newsheet->setCellValue(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1), $value_all);
                            $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray($style_vertical_middle);
                            if(isset($font_info->bl) && $font_info->bl){
                                $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->setBold(true);
                            }
                            if(isset($font_info->it) && $font_info->it){
                                $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->setItalic(true);
                            }
                            if(isset($font_info->fc) && $font_info->fc && isset($json_all[$i]->celldata[$j]->v->fc) && $json_all[$i]->celldata[$j]->v->fc){
                                $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->getColor()->setRGB(substr($json_all[$i]->celldata[$j]->v->fc, 1));
                            }
                            if(isset($font_info->fs) && $font_info->fs){
                                $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->getFont()->setSize($font_info->fs);
                            }
                            if(isset($font_info->ff) && $font_info->ff){
                                $newsheet->getStyle(convert_alphabet($json_all[$i]->celldata[$j]->c)[0] . ((int) $json_all[$i]->celldata[$j]->r + 1))->applyFromArray(array("font" => array("name" => $font_info->ff)));
                            }
                        }
                    }
                }
            }
        }
        // $newsheet->freezePane('A10');
        // Config
        if(isset($json_all[$i]->config)){
            if(isset($json_all[$i]->config->borderInfo)){
                for($j = 0; $j < sizeof($json_all[$i]->config->borderInfo); $j++){
                    $borderInfo = $json_all[$i]->config->borderInfo[$j];
                    if(isset($borderInfo->color) && $borderInfo->color){
                        $border_style = PHPExcel_Style_Border::BORDER_THIN;
                        if($borderInfo->color == "#000"){
                            $borderInfo->color = "#000000";
                        }
                        if($borderInfo->style == "1"){
                            $border_style = PHPExcel_Style_Border::BORDER_THIN;
                        }
                        if($borderInfo->style == "3"){
                            $border_style = PHPExcel_Style_Border::BORDER_DOTTED;
                        }
                        if($borderInfo->style == "4"){
                            $border_style = PHPExcel_Style_Border::BORDER_DASHED;
                        }
                        if($borderInfo->style == "5"){
                            $border_style = PHPExcel_Style_Border::BORDER_DASHDOT;
                        }
                        if($borderInfo->style == "6"){
                            $border_style = PHPExcel_Style_Border::BORDER_DASHDOTDOT;
                        }
                        if($borderInfo->style == "8"){
                            $border_style = PHPExcel_Style_Border::BORDER_MEDIUM;
                        }
                        if($borderInfo->style == "9"){
                            $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHED;
                        }
                        if($borderInfo->style == "10"){
                            $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;
                        }
                        if($borderInfo->style == "11"){
                            $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT;
                        }
                        if($borderInfo->style == "13"){
                            $border_style = PHPExcel_Style_Border::BORDER_THICK;
                        }
                        $style_border_top = array(
                            'borders' => array(
                                'top' => array(
                                    'style' => $border_style,
                                    'color' => array('rgb' => substr($borderInfo->color, 1))
                                )
                            )
                        );
                        $style_border_bottom = array(
                            'borders' => array(
                                'bottom' => array(
                                    'style' => $border_style,
                                    'color' => array('rgb' => substr($borderInfo->color, 1))
                                )
                            )
                        );
                        $style_border_right = array(
                            'borders' => array(
                                'right' => array(
                                    'style' => $border_style,
                                    'color' => array('rgb' => substr($borderInfo->color, 1))
                                )
                            )
                        );
                        $style_border_left = array(
                            'borders' => array(
                                'left' => array(
                                    'style' => $border_style,
                                    'color' => array('rgb' => substr($borderInfo->color, 1))
                                )
                            )
                        );
                    }
                    if($borderInfo->rangeType == "range"){
                        if($borderInfo->borderType == "border-all"){
                            $column_info = $borderInfo->range[0];
                            for($k = $column_info->row[0]; $k <= $column_info->row[1]; $k++){
                                for($h = $column_info->column[0]; $h <= $column_info->column[1]; $h++){
                                    $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_top);
                                    $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                    $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                    $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                }
                            }
                        }
                        if($borderInfo->borderType == "border-right"){
                            $column_info = $borderInfo->range[0];
                            $column_active = $column_info->column[1];
                            for($k = $column_info->row[0]; $k <= $column_info->row[1]; $k++){
                                $newsheet->getStyle(convert_alphabet($column_active)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                            }
                        }
                        if($borderInfo->borderType == "border-left"){
                            $column_info = $borderInfo->range[0];
                            $column_active = $column_info->column[0];
                            for($k = $column_info->row[0]; $k <= $column_info->row[1]; $k++){
                                $newsheet->getStyle(convert_alphabet($column_active)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                            }
                        }
                        if($borderInfo->borderType == "border-top"){
                            $column_info = $borderInfo->range[0];
                            $row_active = $column_info->row[0];
                            for($h = $column_info->column[0]; $h <= $column_info->column[1]; $h++){
                                $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$row_active + 1))->applyFromArray($style_border_top);
                            }
                        }
                        if($borderInfo->borderType == "border-bottom"){
                            $column_info = $borderInfo->range[0];
                            $row_active = $column_info->row[1];
                            for($h = $column_info->column[0]; $h <= $column_info->column[1]; $h++){
                                $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$row_active + 1))->applyFromArray($style_border_bottom);
                            }
                        }
                        if($borderInfo->borderType == "border-inside"){
                            $column_info = $borderInfo->range[0];
                            for($k = $column_info->row[0]; $k <= $column_info->row[1]; $k++){
                                for($h = $column_info->column[0]; $h <= $column_info->column[1]; $h++){
                                    if($k == $column_info->row[0]){
                                        if($h == $column_info->column[0]){
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                        }
                                        else if($h > $column_info->column[0] && $h < $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                        }
                                        else if($h == $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                        }
                                    }
                                    else if($k > $column_info->row[0] && $k < $column_info->row[1]){
                                        if($h == $column_info->column[0]){
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                        }
                                        else if($h > $column_info->column[0] && $h < $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                        }
                                        else if($h == $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                        }
                                    }
                                    else if($k == $column_info->row[1]){
                                        if($h == $column_info->column[0]){
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_top);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                        }
                                        else if($h > $column_info->column[0] && $h < $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_top);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                        }
                                        else if($h == $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_top);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                        }
                                    }
                                }
                            }
                        }
                        if($borderInfo->borderType == "border-outside"){
                            $column_info = $borderInfo->range[0];
                            for($k = $column_info->row[0]; $k <= $column_info->row[1]; $k++){
                                for($h = $column_info->column[0]; $h <= $column_info->column[1]; $h++){
                                    if($k == $column_info->row[0]){
                                        if($h == $column_info->column[0]){
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_top);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                        }
                                        if($h > $column_info->column[0] && $h < $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_top);
                                        }
                                        if($h == $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_top);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                        }
                                    }
                                    if($k > $column_info->row[0] && $k < $column_info->row[1]){
                                        if($h == $column_info->column[0]){
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                        }
                                        if($h == $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                        }
                                    }
                                    if($k == $column_info->row[1]){
                                        if($h == $column_info->column[0]){
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                        }
                                        if($h > $column_info->column[0] && $h < $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                        }
                                        if($h == $column_info->column[1]) {
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                            $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                        }
                                    }
                                }
                            }
                        }
                        if($borderInfo->borderType == "border-horizontal"){
                            $column_info = $borderInfo->range[0];
                            for($k = $column_info->row[0]; $k <= $column_info->row[1]; $k++){
                                for($h = $column_info->column[0]; $h <= $column_info->column[1]; $h++){
                                    if($k >= $column_info->row[0] && $k < $column_info->row[1]){
                                        $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                    }
                                    if($k == $column_info->row[0] && $k == $column_info->row[1]){
                                        $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_bottom);
                                    }
                                }
                            }
                        }
                        if($borderInfo->borderType == "border-vertical"){
                            $column_info = $borderInfo->range[0];
                            for($k = $column_info->row[0]; $k <= $column_info->row[1]; $k++){
                                for($h = $column_info->column[0]; $h <= $column_info->column[1]; $h++){
                                    if($h >= $column_info->column[0] && $h < $column_info->column[1]){
                                        $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                    }
                                    if($h == $column_info->column[0] && $h == $column_info->column[1]){
                                        $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                    }
                                }
                            }
                        }
                        if($borderInfo->borderType == "border-none"){
                            $column_info = $borderInfo->range[0];
                            // echo "border-none\n";
                            $style_border = array(
                                'borders' => array(
                                    'allborders' => array(
                                        'style' => PHPExcel_Style_Border::BORDER_NONE
                                    )
                                )
                            );
                            $style_border_right = array(
                                'borders' => array(
                                    'right' => array(
                                        'style' => PHPExcel_Style_Border::BORDER_NONE
                                    )
                                )
                            );
                            $style_border_left = array(
                                'borders' => array(
                                    'left' => array(
                                        'style' => PHPExcel_Style_Border::BORDER_NONE
                                    )
                                )
                            );
                            $style_border_top = array(
                                'borders' => array(
                                    'top' => array(
                                        'style' => PHPExcel_Style_Border::BORDER_NONE
                                    )
                                )
                            );
                            $style_border_bottom = array(
                                'borders' => array(
                                    'bottom' => array(
                                        'style' => PHPExcel_Style_Border::BORDER_NONE
                                    )
                                )
                            );
                            for($k = $column_info->row[0]; $k <= $column_info->row[1]; $k++){
                                for($h = $column_info->column[0]; $h <= $column_info->column[1]; $h++){
                                    // echo "MASUK BORDER NONE ".convert_alphabet($h)[0] . ((int)$k + 1).".\n";
                                    $newsheet->getStyle(convert_alphabet($h)[0] . ((int)$k + 1))->applyFromArray($style_border);
                                    $newsheet->getStyle(convert_alphabet($h - 1)[0] . ((int)$k + 1))->applyFromArray($style_border_right);
                                    $newsheet->getStyle(convert_alphabet($h + 1)[0] . ((int)$k + 1))->applyFromArray($style_border_left);
                                    $newsheet->getStyle(convert_alphabet($h)[0] . (((int)$k + 1) - 1))->applyFromArray($style_border_bottom);
                                    $newsheet->getStyle(convert_alphabet($h)[0] . (((int)$k + 1) + 1))->applyFromArray($style_border_top);
                                }
                            }
                        }
                    }
                    if($borderInfo->rangeType == "cell"){
                        if(is_object($borderInfo->value)){
                            if(isset($borderInfo->value->t)){
                                $border_style = PHPExcel_Style_Border::BORDER_THIN;
                                if($borderInfo->value->t->color == "#000"){
                                    $borderInfo->value->t->color = "#000000";
                                }
                                if($borderInfo->value->t->style == "1"){
                                    $border_style = PHPExcel_Style_Border::BORDER_THIN;
                                }
                                if($borderInfo->value->t->style == "3"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DOTTED;
                                }
                                if($borderInfo->value->t->style == "4"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHED;
                                }
                                if($borderInfo->value->t->style == "5"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHDOT;
                                }
                                if($borderInfo->value->t->style == "6"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHDOTDOT;
                                }
                                if($borderInfo->value->t->style == "8"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUM;
                                }
                                if($borderInfo->value->t->style == "9"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHED;
                                }
                                if($borderInfo->value->t->style == "10"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;
                                }
                                if($borderInfo->value->t->style == "11"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT;
                                }
                                if($borderInfo->value->t->style == "13"){
                                    $border_style = PHPExcel_Style_Border::BORDER_THICK;
                                }
                                $style_border_top = array(
                                    'borders' => array(
                                        'top' => array(
                                            'style' => $border_style,
                                            'color' => array('rgb' => substr($borderInfo->value->t->color, 1))
                                        )
                                    )
                                );
                                $newsheet->getStyle(convert_alphabet($borderInfo->value->col_index)[0] . ((int)$borderInfo->value->row_index + 1))->applyFromArray($style_border_top);
                            }
                            if(isset($borderInfo->value->b)){
                                $border_style = PHPExcel_Style_Border::BORDER_THIN;
                                if($borderInfo->value->b->color == "#000"){
                                    $borderInfo->value->b->color = "#000000";
                                }
                                if($borderInfo->value->b->style == "1"){
                                    $border_style = PHPExcel_Style_Border::BORDER_THIN;
                                }
                                if($borderInfo->value->b->style == "3"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DOTTED;
                                }
                                if($borderInfo->value->b->style == "4"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHED;
                                }
                                if($borderInfo->value->b->style == "5"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHDOT;
                                }
                                if($borderInfo->value->b->style == "6"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHDOTDOT;
                                }
                                if($borderInfo->value->b->style == "8"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUM;
                                }
                                if($borderInfo->value->b->style == "9"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHED;
                                }
                                if($borderInfo->value->b->style == "10"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;
                                }
                                if($borderInfo->value->b->style == "11"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT;
                                }
                                if($borderInfo->value->b->style == "13"){
                                    $border_style = PHPExcel_Style_Border::BORDER_THICK;
                                }
                                $style_border_bottom = array(
                                    'borders' => array(
                                        'bottom' => array(
                                            'style' => $border_style,
                                            'color' => array('rgb' => substr($borderInfo->value->b->color, 1))
                                        )
                                    )
                                );
                                $newsheet->getStyle(convert_alphabet($borderInfo->value->col_index)[0] . ((int)$borderInfo->value->row_index + 1))->applyFromArray($style_border_bottom);
                            }
                            if(isset($borderInfo->value->r)){
                                $border_style = PHPExcel_Style_Border::BORDER_THIN;
                                if($borderInfo->value->r->color == "#000"){
                                    $borderInfo->value->r->color = "#000000";
                                }
                                if($borderInfo->value->r->style == "1"){
                                    $border_style = PHPExcel_Style_Border::BORDER_THIN;
                                }
                                if($borderInfo->value->r->style == "3"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DOTTED;
                                }
                                if($borderInfo->value->r->style == "4"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHED;
                                }
                                if($borderInfo->value->r->style == "5"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHDOT;
                                }
                                if($borderInfo->value->r->style == "6"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHDOTDOT;
                                }
                                if($borderInfo->value->r->style == "8"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUM;
                                }
                                if($borderInfo->value->r->style == "9"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHED;
                                }
                                if($borderInfo->value->r->style == "10"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;
                                }
                                if($borderInfo->value->r->style == "11"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT;
                                }
                                if($borderInfo->value->r->style == "13"){
                                    $border_style = PHPExcel_Style_Border::BORDER_THICK;
                                }
                                $style_border_right = array(
                                    'borders' => array(
                                        'right' => array(
                                            'style' => $border_style,
                                            'color' => array('rgb' => substr($borderInfo->value->r->color, 1))
                                        )
                                    )
                                );
                                $newsheet->getStyle(convert_alphabet($borderInfo->value->col_index)[0] . ((int)$borderInfo->value->row_index + 1))->applyFromArray($style_border_right);
                            }
                            if(isset($borderInfo->value->l)){
                                $border_style = PHPExcel_Style_Border::BORDER_THIN;
                                if($borderInfo->value->l->color == "#000"){
                                    $borderInfo->value->l->color = "#000000";
                                }
                                if($borderInfo->value->l->style == "1"){
                                    $border_style = PHPExcel_Style_Border::BORDER_THIN;
                                }
                                if($borderInfo->value->l->style == "3"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DOTTED;
                                }
                                if($borderInfo->value->l->style == "4"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHED;
                                }
                                if($borderInfo->value->l->style == "5"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHDOT;
                                }
                                if($borderInfo->value->l->style == "6"){
                                    $border_style = PHPExcel_Style_Border::BORDER_DASHDOTDOT;
                                }
                                if($borderInfo->value->l->style == "8"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUM;
                                }
                                if($borderInfo->value->l->style == "9"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHED;
                                }
                                if($borderInfo->value->l->style == "10"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;
                                }
                                if($borderInfo->value->l->style == "11"){
                                    $border_style = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT;
                                }
                                if($borderInfo->value->l->style == "13"){
                                    $border_style = PHPExcel_Style_Border::BORDER_THICK;
                                }
                                $style_border_left = array(
                                    'borders' => array(
                                        'left' => array(
                                            'style' => $border_style,
                                            'color' => array('rgb' => substr($borderInfo->value->l->color, 1))
                                        )
                                    )
                                );
                                $newsheet->getStyle(convert_alphabet($borderInfo->value->col_index)[0] . ((int)$borderInfo->value->row_index + 1))->applyFromArray($style_border_left);
                            }
                        }
                    }
                }
            }
            if(isset($json_all[$i]->config->merge)){
                $array_merge = (array) $json_all[$i]->config->merge;
                $array_merge_keys = array_keys($array_merge);
                for($j = 0; $j < sizeof($array_merge_keys); $j++){
                    $merge_info = $array_merge[$array_merge_keys[$j]];
                    $newsheet->mergeCells(convert_alphabet($merge_info->c)[0].($merge_info->r + 1).":".convert_alphabet(($merge_info->c + ($merge_info->cs - 1)))[0].(($merge_info->r + 1) + ($merge_info->rs - 1)));
                }
            }
            if(isset($json_all[$i]->config->columnlen)){
                $columnlen_info = (array) $json_all[$i]->config->columnlen;
                $columnlen_info_keys = array_keys($columnlen_info);
                for($j = 0; $j < sizeof($columnlen_info_keys); $j++){
                    $newsheet->getColumnDimension(convert_alphabet($columnlen_info_keys[$j])[0])->setWidth(ceil($columnlen_info[$columnlen_info_keys[$j]] / 8));
                }
            }
            if(isset($json_all[$i]->config->rowlen)){
                $rowlen_info = (array) $json_all[$i]->config->rowlen;
                $rowlen_info_keys = array_keys($rowlen_info);
                for($j = 0; $j < sizeof($rowlen_info_keys); $j++){
                    $newsheet->getRowDimension(($rowlen_info_keys[$j] + 1))->setRowHeight(ceil($rowlen_info[$rowlen_info_keys[$j]]));
                }
            }
        }
        if(isset($json_all[$i]->frozen)){
            $frozenInfo = $json_all[$i]->frozen;
            if($frozenInfo->type == "rangeColumn"){
                $newsheet->freezePane(convert_alphabet(($frozenInfo->range->column_focus + 1))[0].'1');
            }
            if($frozenInfo->type == "rangeRow"){
                $newsheet->freezePane('A'.($frozenInfo->range->row_focus + 2));
            }
            if($frozenInfo->type == "rangeBoth"){
                $newsheet->freezePane(convert_alphabet(($frozenInfo->range->column_focus + 1))[0].($frozenInfo->range->row_focus + 2));
            }
        }
    }
    $excel->setActiveSheetIndex(0);
    ob_start();
    $excel_name = "YKKBI-LUCKYSHEET".date("YmdHis").round(microtime(true) * 1000).".xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="'.$excel_name.'"'); // Set nama file excel nya
    header('Cache-Control: max-age=0');
    $write = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
    $write->save('php://output');
    $excel_content = ob_get_clean();
    file_put_contents(__DIR__."/phpexcel-output/" . $excel_name, $excel_content);
} else {
    echo "<html>\n";
    echo "<head>\n";
    echo "<title>Wrong Method.</title>\n";
    echo "<style type='text/css'>html, body {font-family: consolas, monsopace;}</style>\n";
    echo "</head>\n";
    echo "<body>\n";
    echo "Wrong Method.\n";
    echo "</body>\n";
    echo "</html>\n";
}

