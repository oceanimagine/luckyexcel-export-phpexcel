<?php

/* 
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Scripting/EmptyPHP.php to edit this template
 */

file_put_contents("POST_RAW_2", print_r(json_decode(json_encode($_POST['isi_excel'])), true));