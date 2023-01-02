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


$app->get('/checkuser', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
  
   $user_id = isset($_REQUEST['user_id']) ? $_REQUEST['user_id'] : '';
  $username = "webcontrol"; // Use your username
$password = "webcontrol"; // and your password
$database = "(DESCRIPTION=(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.80)(PORT = 1521)))(CONNECT_DATA=(SID=NYTG)))"; 
$objConnect = oci_connect($username, $password, $database, 'AL32UTF8');
  echo $strSQL = "SELECT * FROM CTL_USER where user_id =  '".$user_id."'";
  $objParse = oci_parse($objConnect, $strSQL);
  oci_execute ($objParse,OCI_DEFAULT);
 while($objResult = oci_fetch_array($objParse,OCI_BOTH)){
      $arr1= array(
        'Status'  => true,
        'member' => $objResult['USER_ID'],
        'NAME' => $objResult['USER_FNAME'],
        'LNAME' => $objResult['USER_LNAME'],
        'LANG' => $objResult['DEFAULT_LANGUAGE']
      );
      $array2[] = $arr1;
  }
   
   $data = json_encode($array2);
       
  
  $response->getBody()->write($data); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า */
});




$app->run(); //สั่งระบบให้ทำงาน

?>