<?php
ini_set("display_errors", 0);
error_reporting(0);
if(!function_exists("mysql_connect")){
    $connect_all;
    $host_all;
    $user_all;
    $pass_all;
    $data_all;

    function mysql_connect($host, $user, $pass){
        global $host_all;
        global $user_all;
        global $pass_all;
        $host_all = $host;
        $user_all = $user;
        $pass_all = $pass;
        return true;
    }

    function mysql_select_db($db_name){
        global $connect_all;
        global $host_all;
        global $user_all;
        global $pass_all;
        $connect_all = mysqli_connect($host_all, $user_all, $pass_all, $db_name);
        if($connect_all){
            return true;
        } else {
            return false;
        }
    }

    function mysql_close(){
        global $connect_all;
        mysqli_close($connect_all);
    }

    function mysql_query($sql){
        // echo $sql;
        global $connect_all;
        // print_r($connect_all);
        // $query = mysqli_query($connect_all, $sql);;
        // echo mysqli_num_rows($query);
        return mysqli_query($connect_all, $sql);
    }

    function mysql_fetch_row($result){
        return mysqli_fetch_row($result);
    }

    function mysql_fetch_array($result){
        return mysqli_fetch_array($result);
    }

    function mysql_num_rows($result){
        return mysqli_num_rows($result);
    }

    function mysql_fetch_assoc($result){
        return mysqli_fetch_assoc($result);
    }
    function mysql_real_escape_string($param){
        global $connect_all;
        return mysqli_real_escape_string($connect_all, $param);
    }

    function split($delimeter, $string){
        return explode($delimeter, $string);
    }
    function mysql_affected_rows(){
        global $connect_all;
        return mysqli_affected_rows($connect_all);
    }
}
include_once "config/config.php";
$connect = mysql_connect(dbHost, dbUser, dbPass);
mysql_select_db(dbName);
require_once __DIR__.'/phpexcel/PHPExcel.php';

$excel = new PHPExcel();
$excel->getProperties()->setCreator('Gemah Ripah Deviden')
        ->setLastModifiedBy('Gemah Ripah Deviden')
        ->setTitle("Gemah Ripah Deviden")
        ->setSubject("Gemah Ripah Deviden")
        ->setDescription("Gemah Ripah Deviden")
        ->setKeywords("Gemah Ripah Deviden");

$style_col = array(
    'font' => array('bold' => true),
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
    ),
    'borders' => array(
        'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
        'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
        'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
        'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN)
    )
);

$style_col_body = array(
    'alignment' => array(
        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
    ),
    'borders' => array(
        'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
        'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
        'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
        'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN)
    )
);

$tahun = "";
if (defined('STDIN')) {
    $tahun = $argv[1];
}
if($tahun == ""){
    echo "\n";
    echo "Please Define Year Parameter.\n";
    exit();
}

$_GET['year'] = $tahun;
$year = isset($_GET['year']) ? $_GET['year'] : "";
$query_parameter = mysql_query("select shu, balas_jasa, status_proses from t_new_parameter_deviden where tahun = '".$year."'");
$shu = "";
$balas_jasa = "";
if(mysql_num_rows($query_parameter) > 0){
    $hasil_parameter = mysql_fetch_array($query_parameter);
    $shu = $hasil_parameter['shu'];
    $balas_jasa = $hasil_parameter['balas_jasa'];
}

$excel->setActiveSheetIndex(0)->setCellValue('B1', "80% SHU (L/R)");
$excel->setActiveSheetIndex(0)->setCellValue('C1', $shu);
$excel->setActiveSheetIndex(0)->setCellValue('B2', "Balas_Jasa");
$excel->setActiveSheetIndex(0)->setCellValue('B3', "Deviden_Dibagikan");
$excel->setActiveSheetIndex(0)->setCellValue('B4', "Jml_Total_Bln_Saham");
$excel->setActiveSheetIndex(0)->setCellValue('B5', "Deviden_Per_Lembar_Saham");
$excel->setActiveSheetIndex(0)->setCellValue('B6', "%Balas Jasa");
$excel->setActiveSheetIndex(0)->setCellValue('C6', $balas_jasa);
$excel->setActiveSheetIndex(0)->setCellValue('B7', "%Avg_Deviden_VS_Simpanan");

$excel->setActiveSheetIndex(0)->setCellValue('AH8', $year);
$excel->getActiveSheet()->mergeCells('AH8:AO8');
$excel->getActiveSheet()->getStyle('AH8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$excel->setActiveSheetIndex(0)->getStyle("AH8:AO8")->applyFromArray($style_col_body);

$excel->setActiveSheetIndex(0)->setCellValue("A9", "Nomor Anggota");
$excel->setActiveSheetIndex(0)->setCellValue("B9", "Nama Anggota");
$excel->setActiveSheetIndex(0)->setCellValue("C9", "Payroll");
$excel->setActiveSheetIndex(0)->setCellValue("D9", "Unit Simpin");
$excel->setActiveSheetIndex(0)->setCellValue("E9", "Unit Group Simpin");
$excel->setActiveSheetIndex(0)->setCellValue("F9", "Group Description SAP");
$excel->setActiveSheetIndex(0)->setCellValue("G9", "Division Description SAP");
$excel->setActiveSheetIndex(0)->setCellValue("H9", "Simpanan Desember " . ($year - 1));
$excel->setActiveSheetIndex(0)->setCellValue("I9", "Bulan Saham");
$excel->setActiveSheetIndex(0)->setCellValue("J9", "BS Januari");
$excel->setActiveSheetIndex(0)->setCellValue("K9", "BS Februari");
$excel->setActiveSheetIndex(0)->setCellValue("L9", "BS Maret");
$excel->setActiveSheetIndex(0)->setCellValue("M9", "BS April");
$excel->setActiveSheetIndex(0)->setCellValue("N9", "BS Mei");
$excel->setActiveSheetIndex(0)->setCellValue("O9", "BS Juni");
$excel->setActiveSheetIndex(0)->setCellValue("P9", "BS Juli");
$excel->setActiveSheetIndex(0)->setCellValue("Q9", "BS Agustus");
$excel->setActiveSheetIndex(0)->setCellValue("R9", "BS September");
$excel->setActiveSheetIndex(0)->setCellValue("S9", "BS Oktober");
$excel->setActiveSheetIndex(0)->setCellValue("T9", "BS November");
$excel->setActiveSheetIndex(0)->setCellValue("U9", "BS Jumlah");
$excel->setActiveSheetIndex(0)->setCellValue("V9", "Simpanan Januari " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("W9", "Simpanan Februari " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("X9", "Simpanan Maret " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("Y9", "Simpanan April " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("Z9", "Simpanan Mei " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AA9", "Simpanan Juni " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AB9", "Simpanan Juli " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AC9", "Simpanan Agustus " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AD9", "Simpanan September " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AE9", "Simpanan Oktober " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AF9", "Simpanan November " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AG9", "Simpanan Desember " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AH9", "Sisa Pinjaman " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AI9", "Bunga " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AJ9", "Balas Jasa " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AK9", "Deviden " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AL9", "Balas Jasa + Deviden " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AM9", "Deviden Ditransfer " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AN9", "Keterangan " . $year);
$excel->setActiveSheetIndex(0)->setCellValue("AO9", "Ikut Deviden");

$excel->setActiveSheetIndex(0)->getStyle("A9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("B9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("C9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("D9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("E9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("F9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("G9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("H9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("I9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("J9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("K9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("L9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("M9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("N9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("O9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("P9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("Q9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("R9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("S9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("T9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("U9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("V9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("W9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("X9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("Y9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("X9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("Z9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AA9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AB9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AC9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AD9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AE9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AF9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AG9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AH9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AI9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AJ9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AK9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AL9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AM9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AN9")->applyFromArray($style_col_body);
$excel->setActiveSheetIndex(0)->getStyle("AO9")->applyFromArray($style_col_body);

$jumlah_bs = 0;
$jumlah_balas_jasa = 0;
$jumlah_simpanan_desember_tahun_sekarang = 0;
$array_bs = array();
$array_balas_jasa = array();
$query_all_years = mysql_query("SELECT no_anggota,nama,payroll,unit,unit_group FROM `t_new_deviden_isi_excel_bulanan` where periode_tahun = '".$year."' and `periode_bulan` = '12'");
if(mysql_num_rows($query_all_years) > 0){
    $nomor = 10;
    $awal_ = $nomor;
    while($hasil_all_years = mysql_fetch_array($query_all_years)){
        echo $hasil_all_years['no_anggota'] . "\n";
        $excel->setActiveSheetIndex(0)->setCellValue("A" . $nomor, $hasil_all_years['no_anggota']);
        $excel->setActiveSheetIndex(0)->setCellValue("B" . $nomor, $hasil_all_years['nama']);
        $excel->setActiveSheetIndex(0)->setCellValue("C" . $nomor, $hasil_all_years['payroll']);
        $excel->setActiveSheetIndex(0)->setCellValue("D" . $nomor, $hasil_all_years['unit']);
        $excel->setActiveSheetIndex(0)->setCellValue("E" . $nomor, $hasil_all_years['unit_group']);
        $excel->setActiveSheetIndex(0)->setCellValue("F" . $nomor, "-");
        $excel->setActiveSheetIndex(0)->setCellValue("G" . $nomor, "-");
        $simp_12_tahun_lalu = 0;
        $query_simp_12_tahun_lalu = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '12' and 
            periode_tahun = '".($year - 1)."'
        ");
        if(mysql_num_rows($query_simp_12_tahun_lalu) > 0){
            $hasil_simp_12_tahun_lalu = mysql_fetch_array($query_simp_12_tahun_lalu);
            $simp_12_tahun_lalu = $hasil_simp_12_tahun_lalu['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("H" . $nomor, $simp_12_tahun_lalu);
        $excel->setActiveSheetIndex(0)->setCellValue("I" . $nomor, "=H".$nomor."/2000*12");
        
        $simp_01 = "0";
        $query_simp_01 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '01' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_01) > 0){
            $hasil_simp_01 = mysql_fetch_array($query_simp_01);
            $simp_01 = $hasil_simp_01['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("J" . $nomor, "=(V".$nomor."-H".$nomor.")/2000*11");
        
        $simp_02 = "0";
        $query_simp_02 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '02' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_02) > 0){
            $hasil_simp_02 = mysql_fetch_array($query_simp_02);
            $simp_02 = $hasil_simp_02['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("K" . $nomor, "=(W".$nomor."-V".$nomor.")/2000*10");
        
        $simp_03 = "0";
        $query_simp_03 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '03' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_03) > 0){
            $hasil_simp_03 = mysql_fetch_array($query_simp_03);
            $simp_03 = $hasil_simp_03['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("L" . $nomor, "=(X".$nomor."-W".$nomor.")/2000*9");
        
        $simp_04 = "0";
        $query_simp_04 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '04' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_04) > 0){
            $hasil_simp_04 = mysql_fetch_array($query_simp_04);
            $simp_04 = $hasil_simp_04['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("M" . $nomor, "=(Y".$nomor."-X".$nomor.")/2000*8");
        
        $simp_05 = "0";
        $query_simp_05 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '05' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_05) > 0){
            $hasil_simp_05 = mysql_fetch_array($query_simp_05);
            $simp_05 = $hasil_simp_05['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("N" . $nomor, "=(Z".$nomor."-Y".$nomor.")/2000*7");
        
        $simp_06 = "0";
        $query_simp_06 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '06' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_06) > 0){
            $hasil_simp_06 = mysql_fetch_array($query_simp_06);
            $simp_06 = $hasil_simp_06['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("O" . $nomor, "=(AA".$nomor."-Z".$nomor.")/2000*6");
        
        
        $simp_07 = "0";
        $query_simp_07 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '07' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_07) > 0){
            $hasil_simp_07 = mysql_fetch_array($query_simp_07);
            $simp_07 = $hasil_simp_07['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("P" . $nomor, "=(AB".$nomor."-AA".$nomor.")/2000*5");
        
        
        $simp_08 = "0";
        $query_simp_08 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '08' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_08) > 0){
            $hasil_simp_08 = mysql_fetch_array($query_simp_08);
            $simp_08 = $hasil_simp_08['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("Q" . $nomor, "=(AC".$nomor."-AB".$nomor.")/2000*4");
        
        
        $simp_09 = "0";
        $query_simp_09 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '09' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_09) > 0){
            $hasil_simp_09 = mysql_fetch_array($query_simp_09);
            $simp_09 = $hasil_simp_09['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("R" . $nomor, "=(AD".$nomor."-AC".$nomor.")/2000*3");
        
        
        $simp_10 = "0";
        $query_simp_10 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '10' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_10) > 0){
            $hasil_simp_10 = mysql_fetch_array($query_simp_10);
            $simp_10 = $hasil_simp_10['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("S" . $nomor, "=(AE".$nomor."-AD".$nomor.")/2000*2");
        
        $simp_11 = "0";
        $query_simp_11 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '11' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_11) > 0){
            $hasil_simp_11 = mysql_fetch_array($query_simp_11);
            $simp_11 = $hasil_simp_11['total_simpanan'];
        }
        
        $excel->setActiveSheetIndex(0)->setCellValue("T" . $nomor, "=(AF".$nomor."-AE".$nomor.")/2000*1");
        $excel->setActiveSheetIndex(0)->setCellValue("U" . $nomor, "=SUM(I".$nomor.":T".$nomor.")");
        
        $excel->setActiveSheetIndex(0)->setCellValue("V" . $nomor, $simp_01);
        $excel->setActiveSheetIndex(0)->setCellValue("W" . $nomor, $simp_02);
        $excel->setActiveSheetIndex(0)->setCellValue("X" . $nomor, $simp_03);
        $excel->setActiveSheetIndex(0)->setCellValue("Y" . $nomor, $simp_04);
        $excel->setActiveSheetIndex(0)->setCellValue("Z" . $nomor, $simp_05);
        $excel->setActiveSheetIndex(0)->setCellValue("AA" . $nomor, $simp_06);
        $excel->setActiveSheetIndex(0)->setCellValue("AB" . $nomor, $simp_07);
        $excel->setActiveSheetIndex(0)->setCellValue("AC" . $nomor, $simp_08);
        $excel->setActiveSheetIndex(0)->setCellValue("AD" . $nomor, $simp_09);
        $excel->setActiveSheetIndex(0)->setCellValue("AE" . $nomor, $simp_10);
        $excel->setActiveSheetIndex(0)->setCellValue("AF" . $nomor, $simp_11);
        
        
        $simp_12 = "0";
        $query_simp_12 = mysql_query("
            select total_simpanan from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' 
            and periode_bulan = '12' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_simp_12) > 0){
            $hasil_simp_12 = mysql_fetch_array($query_simp_12);
            $simp_12 = $hasil_simp_12['total_simpanan'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("AG" . $nomor, $simp_12);
        
        $jumlah_simpanan_desember_tahun_sekarang = $jumlah_simpanan_desember_tahun_sekarang + $simp_12;
        
        $sisa_pinjaman = "0";
        $query_sisa_pinjaman = mysql_query("
            select sisa_pinjaman from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' and
            periode_bulan = '12' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_sisa_pinjaman) > 0){
            $hasil_sisa_pinjaman = mysql_fetch_array($query_sisa_pinjaman);
            $sisa_pinjaman = $hasil_sisa_pinjaman['sisa_pinjaman'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("AH" . $nomor, $sisa_pinjaman);
        
        
        $bunga = "0";
        $query_bunga = mysql_query("
            select jml_bunga_dibayar from t_new_deviden_isi_excel_bulanan where no_anggota = '".$hasil_all_years['no_anggota']."' and
            periode_bulan = '12' and 
            periode_tahun = '".$year."'
        ");
        if(mysql_num_rows($query_bunga) > 0){
            $hasil_bunga = mysql_fetch_array($query_bunga);
            $bunga = $hasil_bunga['jml_bunga_dibayar'];
        }
        $excel->setActiveSheetIndex(0)->setCellValue("AI" . $nomor, $bunga);
        $excel->setActiveSheetIndex(0)->setCellValue("AJ" . $nomor, '=ROUNDUP(AI'.$nomor.'*($C$6/100),-3)');
        $array_balas_jasa[] = round($bunga * ($balas_jasa / 100), -3);
        
        $excel->setActiveSheetIndex(0)->setCellValue("AK" . $nomor, '=ROUNDUP(U'.$nomor.'*$C$5,-3)');
        $excel->setActiveSheetIndex(0)->setCellValue("AL" . $nomor, "=AJ".$nomor."+AK".$nomor);
        $excel->setActiveSheetIndex(0)->setCellValue("AM" . $nomor, "=ROUNDDOWN(AL".$nomor.",-3)");
        $excel->setActiveSheetIndex(0)->setCellValue("AN" . $nomor, "TRANSFER");
        $excel->setActiveSheetIndex(0)->setCellValue("AO" . $nomor, "1");
        
        $jumlah_balas_jasa = $jumlah_balas_jasa + round($bunga * ($balas_jasa / 100), -3);
        
        $excel->setActiveSheetIndex(0)->getStyle("A" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("B" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("C" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("D" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("E" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("F" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("G" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("H" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("I" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("J" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("K" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("L" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("M" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("N" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("O" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("P" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("Q" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("R" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("S" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("T" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("U" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("V" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("W" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("X" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("Y" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("X" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("Z" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AA" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AB" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AC" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AD" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AE" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AF" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AG" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AH" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AI" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AJ" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AK" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AL" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AM" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AN" . $nomor)->applyFromArray($style_col_body);
        $excel->setActiveSheetIndex(0)->getStyle("AO" . $nomor)->applyFromArray($style_col_body);
        if($nomor - 10 > 12000){
            echo "Data Count Cannot Larger Than 12000.\n";
            exit();
            break;
        }
        $nomor++;
    }
}

$excel->setActiveSheetIndex(0)->setCellValue('C4', "=SUM(U".$awal_.":U".$nomor.")");
$excel->setActiveSheetIndex(0)->setCellValue('C2', "=SUM(AJ".$awal_.":AJ".$nomor.")");
$excel->setActiveSheetIndex(0)->setCellValue('C3', "=C1-C2");
$excel->setActiveSheetIndex(0)->setCellValue('C5', "=C3/C4");
$excel->setActiveSheetIndex(0)->setCellValue('C7', "=C3/(SUM(AG".$awal_.":AG".$nomor."))");
$excel->setActiveSheetIndex(0)->setCellValue("AM" . $nomor, "=SUM(AM".$awal_.":AM".($nomor - 1).")");

for ($i = 'A'; $i !=  $excel->getActiveSheet()->getHighestColumn(); $i++) {
    $excel->getActiveSheet()->getColumnDimension($i)->setAutoSize(TRUE);
}
$excel->getActiveSheet()->getColumnDimension($i)->setAutoSize(TRUE);
$excel->getActiveSheet()->freezePane('D10');

$excel->getActiveSheet(0)->setTitle("Laporan Deviden");
ob_start();
$excel_name = "DEVIDEN".date("YmdHis").round(microtime(true) * 1000).".xlsx";
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="'.$excel_name.'"'); // Set nama file excel nya
header('Cache-Control: max-age=0');
$write = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$write->save('php://output');

$excel_content = ob_get_clean();
file_put_contents(__DIR__."/DEVIDEN_EXCEL/" . $excel_name, $excel_content);

mysql_query("update t_new_parameter_deviden set file_deviden = '".$excel_name."', status_proses = '(DONE)' where tahun = '".$year."'");
echo "\n\n";
if(mysql_affected_rows() > 0){
    echo "Update Table Parameter Berhasil.\n";
} else {
    echo "Update Table Parameter Gagal.\n";
}

echo "\n";
echo "Mulai Memasukkan Data ke Tabel Terkait.\n";
$tabel_a_deviden_release = "a_deviden_release";
$tabel_a_deviden = "a_deviden";

$sql = "DELETE FROM " . $tabel_a_deviden_release;
mysql_query($sql);
echo "Hapus data dari table ".$tabel_a_deviden_release.".\n";
$sql = "DELETE FROM ".$tabel_a_deviden." WHERE tahun = '".$year."'";
mysql_query($sql);
$berhasil_masuk = 0;
echo "Hapus data dari tabel ".$tabel_a_deviden." untuk tahun '".$year."'.\n\n";
$inputFileName = __DIR__ . '/DEVIDEN_EXCEL/' . $excel_name;
try{
    $inputFileType  =   PHPExcel_IOFactory::identify($inputFileName);
    $objReader      =   PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel    =   $objReader->load($inputFileName);
}catch(Exception $e){
    die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
}
$sheet = $objPHPExcel->getActiveSheet();
$highestRow = $sheet->getHighestRow(); 
for ($row = 10; $row < $highestRow - 1; $row++){
    $no_anggota = str_replace("'", "", $sheet->getCell('A'.$row)->getValue());
    $balas_jasa = $sheet->getCell('AJ'.$row)->getCalculatedValue();
    $deviden = $sheet->getCell('AK'.$row)->getCalculatedValue();
    $balas_jasa_deviden = $sheet->getCell('AL'.$row)->getCalculatedValue();
    $sql = "
    INSERT INTO ".$tabel_a_deviden." SET 
        no_anggota='".$no_anggota."',
        tahun='".$year."',
        balas_jasa='".$balas_jasa."',
        deviden='".$deviden."', 
        jumlah='".$balas_jasa_deviden."'
    ";
    mysql_query($sql);
    if(mysql_affected_rows() > 0){
        $berhasil_masuk = 1;
        echo "Berhasil Insert Deviden nomor anggota ".$array_anggota_release[$i]." balas jasa ".$balas_jasa.".\n";
    }
}
$sql = "INSERT INTO ".$tabel_a_deviden_release." SET release_date=NOW()";
mysql_query($sql);
if(mysql_affected_rows() > 0){
    echo "Berhasil Insert Data ke tabel ".$tabel_a_deviden_release.".\n";
}
    


// echo $excel_content;
