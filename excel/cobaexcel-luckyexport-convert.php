<?php

function convert_to_xml($filename){
    $content = '';
    if(!$filename || !file_exists($filename)) return false;
    $zip = zip_open($filename);
    if (!$zip || is_numeric($zip)) return false;
    while ($zip_entry = zip_read($zip)) {
            if (zip_entry_open($zip, $zip_entry) == FALSE) continue;
            // if (zip_entry_name($zip_entry) != "word/document.xml") continue;
            $content .= zip_entry_read($zip_entry, zip_entry_filesize($zip_entry));
            zip_entry_close($zip_entry);
    }// end while
    zip_close($zip);
    return $content;
}

file_put_contents(__DIR__."/COBA.xml",convert_to_xml(__DIR__ . "/phpexcel-output/YKKBI-LUCKYSHEET202207131233041657708384610.xlsx"));