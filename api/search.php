<?php
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: access");
header("Access-Control-Allow-Methods: GET,POST");
header("Access-Control-Allow-Credentials: true");
header('Content-Type: application/json;charset=utf-8');
require_once __DIR__ . '/../api/web_config.php';

set_time_limit ( 60000 );

use \Psr\Http\Message\ResponseInterface as Response; // ไลบราลี้สำหรับจัดการคำร้องขอ
use \Psr\Http\Message\ServerRequestInterface as Request; // ไลบราลี้สำหรับจัดการคำตอบกลับ

require './vendor/autoload.php'; // ดึงไฟ์ autoload.php เข้ามา
//include_once './class.oracle.php'; // Class Connect Oracle
include_once './util.php'; // ดึงไฟ์ util.php เข้ามา
include_once './web_config.php';
$app = new \Slim\App; // สร้าง object หลักของระบบ



function ConnectDbAll($_sql){
  $DataRows = array();
  $conn = ConnectedDBSO();
  if(!$conn){
    $_err = oci_error();
    echo $_err;
  }else{
    $objParse = oci_parse($conn,$_sql);
    $objEx = oci_execute($objParse);
    if($objEx){
      $objResult = oci_fetch_all($objParse,$DataRows,null,null, OCI_FETCHSTATEMENT_BY_ROW);
    }else{
      echo "Connect Data Base error";
    }
  }
  return $DataRows;
}

 $app->post('/search', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $search = isset($_REQUEST['search']) ? $_REQUEST['search'] : '';
  
  
 
   $_sql = "SELECT SO_NO_DOC, SO_NO, SO_YEAR, QTY, CUST_NAME, STYLE_REF, 
  SHIPMENT_DATE, GMT_TYPE,FABRIC_DESCRIPTION FABRIC_TYPE
   FROM oe_ME_ORDER WHERE SO_YEAR>='21' AND SO_NO='".$search."'";
  $result_out = new stdClass();
  $DataRows = [];
  $resultArray = array(); //data
  $DataRows = ConnectDbAll($_sql);

  if($DataRows != false){
  foreach ($DataRows as $row){
    array_push($resultArray,$row);
  }
  $result_out->data =  $resultArray;
  $result_out->status = (true);
  
}else{
  $result_out->data =  $resultArray;
  $result_out->status = (false);
}

  $SearchArray = [

  ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});



$app->run(); //สั่งระบบให้ทำงาน

?>