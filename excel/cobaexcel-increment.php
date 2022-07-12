<?php

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

$param_hex = 0;
$param_dec = 0;
$param_oct = 0;
$param_bin = 0;
$param_alp = "A";
echo "<style type='text/css'>html, body {font-family: consolas, monospace; padding: 0px; margin: 0px;}</style>\n";
echo "<table border='1' cellspacing='0' cellpadding='15'>\n";
echo "<thead>\n";
echo "<tr>\n";
echo "<th>HEX</th>\n";
echo "<th>DEC</th>\n";
echo "<th>OCT</th>\n";
echo "<th>BIN</th>\n";
echo "<th>ALP</th>\n";
echo "</tr>\n";
echo "</thead>\n";
echo "<tbody>\n";
for($i = 1; $i <= 1000; $i++){
    echo "<tr>\n";
    echo "<td style='text-align: right;'>".$param_hex."</td>\n";
    echo "<td style='text-align: right;'>".$param_dec."</td>\n";
    echo "<td style='text-align: right;'>".$param_oct."</td>\n";
    echo "<td style='text-align: right;'>".$param_bin."</td>\n";
    echo "<td style='text-align: right;'>".$param_alp."</td>\n";
    echo "</tr>\n";
    $param_hex = do_increment($param_hex, "HEX");
    $param_dec = do_increment($param_dec, "DEC");
    $param_oct = do_increment($param_oct, "OCT");
    $param_bin = do_increment($param_bin, "BIN");
    $param_alp = do_increment($param_alp, "ALP");
}
echo "</tbody>\n";
echo "</table>\n";