<?php
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: access");
header("Access-Control-Allow-Methods: GET,POST");
header("Access-Control-Allow-Credentials: true");
header('Content-Type: application/json;charset=utf-8');

set_time_limit ( 60000 );

use \Psr\Http\Message\ResponseInterface as Response; // ไลบราลี้สำหรับจัดการคำร้องขอ
use \Psr\Http\Message\ServerRequestInterface as Request; // ไลบราลี้สำหรับจัดการคำตอบกลับ

require './vendor/autoload.php'; // ดึงไฟ์ autoload.php เข้ามา
//include_once './class.oracle.php'; // Class Connect Oracle
include_once './util.php'; // ดึงไฟ์ util.php เข้ามา
include_once './web_config.php';
$app = new \Slim\App; // สร้าง object หลักของระบบ




$app->post('/checkso', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  /*   $SO_NO = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO'] : '';
   $SO_YEAR = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : ''; */
 
    $SO_NO = '2536';
   $SO_YEAR = '21';  
  $username = "NYGM"; // Use your username
 $password = "NYGM"; // and your password
 $database = "(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.83)(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = NYTG)))"; 
 $objConnect = oci_connect($username, $password, $database, 'AL32UTF8');
 $strSQL = "SELECT DISTINCT CPART_NO,USAGE CPART_NAME FROM OE_FABRIC_BILL_HIS
 WHERE SO_YEAR>='".$SO_YEAR."' AND SO_NO='".$SO_NO."'
 ORDER BY CPART_NO";
  $objParse = oci_parse($objConnect, $strSQL);
  oci_execute ($objParse,OCI_DEFAULT);
 while($objResult = oci_fetch_array($objParse,OCI_BOTH)){
      $arr1= array(
        'Status'  => true,
        'Cpartno' => $objResult['CPART_NO'],
        'Cpartname' => $objResult['CPART_NAME']
      );
      $array2[] = $arr1;
  }
    
  $data = json_encode($array2); 
      
  
  $response->getBody()->write($data); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});


$app->post('/checkgm', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $SO_NO = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $CPART_NO = isset($_REQUEST['CPART_NO']) ? $_REQUEST['CPART_NO'] : '';
  
 $username = "NYGM"; // Use your username
$password = "NYGM"; // and your password
$database = "(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.83)(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = NYTG)))"; 
$objConnect = oci_connect($username, $password, $database, 'AL32UTF8');
$strSQL = "SELECT DISTINCT GM_NO FROM OE_CUT_PLAN_GM08
WHERE SO_YEAR>=21 AND SO_NO='".$SO_NO."' AND CPART_NO='".$CPART_NO."'
ORDER  BY GM_NO";
 $objParse = oci_parse($objConnect, $strSQL);
 oci_execute ($objParse,OCI_DEFAULT);
while($objResult = oci_fetch_array($objParse,OCI_BOTH)){
     $arr1= array(
       'Status'  => true,
       'GM_NO' => $objResult['GM_NO']
     );
     $array2[] = $arr1;
 }
   
 $data = json_encode($array2); 
     
 
 $response->getBody()->write($data); // สงคำตอบกลับ

 return $response; // ส่งคำตอบกลับร้า
});
$app->post('/soinfo', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $SO_NO = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $SO_YEAR = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';

  
 $username = "NYGM"; // Use your username
$password = "NYGM"; // and your password
$database = "(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.83)(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = NYTG)))"; 
$objConnect = oci_connect($username, $password, $database, 'AL32UTF8');
$strSQL = "SELECT SO_NO_DOC, SO_NO, SO_YEAR, QTY, CUST_NAME, STYLE_REF, 
SHIPMENT_DATE, GMT_TYPE,FABRIC_DESCRIPTION FABRIC_TYPE
 FROM oe_ME_ORDER WHERE SO_YEAR>='".$SO_YEAR."' AND SO_NO='".$SO_NO."'";
 $objParse = oci_parse($objConnect, $strSQL);
 oci_execute ($objParse,OCI_DEFAULT);
while($objResult = oci_fetch_array($objParse,OCI_BOTH)){
     $arr1= array(
       'Status'  => true,
       'SO_NO_DOC' => $objResult['SO_NO_DOC'],
       'STYLE_REF' => $objResult['STYLE_REF'],
       'CUS_NAME' => $objResult['CUST_NAME']
     );
     $array2[] = $arr1;
 }
   
 $data = json_encode($array2); 
     
 
 $response->getBody()->write($data); // สงคำตอบกลับ

 return $response; // ส่งคำตอบกลับร้า
});



$app->run(); //สั่งระบบให้ทำงาน

?>