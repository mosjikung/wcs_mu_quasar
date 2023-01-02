<?php

error_reporting(E_ALL);
ini_set('display_errors', 'On');

$username = "webcontrol"; // Use your username
  $password = "webcontrol"; // and your password
  $database = "(DESCRIPTION=(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.80)(PORT = 1521)))(CONNECT_DATA=(SID=NYTG)))"; 
  $objConnect = oci_connect($username, $password, $database, 'AL32UTF8');
    $strSQL = "SELECT * FROM CTL_USER";
    $objParse = oci_parse($objConnect, $strSQL);
    oci_execute ($objParse,OCI_DEFAULT);
    while($objResult = oci_fetch_array($objParse,OCI_BOTH)){
       echo $objResult['USER_FNAME'];
        
    }
