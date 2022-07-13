<?php

function convertSerialDate($date){
    $timestamp = ($date - 25569) * 86400;
    return date("m/d/y",$timestamp);
}

print convertSerialDate(43466);