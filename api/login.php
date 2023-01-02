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

define("LogFolder", "LogPdaMoveloadtor");

$app = new \Slim\App; // สร้าง object หลักของระบบ

function OracleSelectAll($_baselocation, $_query) //ดึง function มากจาก web_config.php
{
    $conn = GetConnectByLocation($_baselocation); //จากการเลือก location
    $DataRows = array();

    if (!$conn) {
        $_err = oci_error();
        $_res = $_err['message'];
        wh_log_path(LogFolder, $_res);
    }
    try {
        $stid = oci_parse($conn, $_query);
        $_res = oci_execute($stid);
        if ($_res) {
            wh_log_path(LogFolder, 'true');
            $nrows = oci_fetch_all($stid, $DataRows, null, null, OCI_FETCHSTATEMENT_BY_ROW);
        } else {
            wh_log_path(LogFolder, 'false');
            $DataRows = false;
        }
    } catch (Throwable $e) {
        // Executed only in PHP 7, will not match in PHP 5.x
        $_res = $e;
        //wh_log($_res);
        wh_log_path(LogFolder, $_res);
    } catch (Exception $e) {
        // Executed only in PHP 5.x, will not be reached in PHP 7
        $_res = $e;
        //wh_log($_res);
        wh_log_path(LogFolder, $_res);
    } finally {
        oci_close($conn);
    }
    return $DataRows;
}

$app->post('/test', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    
    $username = "webcontrol"; // Use your username
  $password = "webcontrol"; // and your password
  $database = "(DESCRIPTION=(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.80)(PORT = 1521)))(CONNECT_DATA=(SID=NYTG)))"; 
  $objConnect = oci_connect($username, $password, $database, 'AL32UTF8');
    $strSQL = "SELECT * FROM CTL_USER";
    $objParse = oci_parse($objConnect, $strSQL);
    oci_execute ($objParse,OCI_DEFAULT);
    while($objResult = oci_fetch_array($objParse,OCI_BOTH)){
      array_push($objResult['USER_FNAME']);
        
    }
    $response->getBody()->write(json_encode($objResult['USER_FNAME'])); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/aut', function (Request $request, Response $response) {


    $UserID = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
    $Password = isset($_REQUEST['password']) ? $_REQUEST['password'] : '';
    $Location = isset($_REQUEST['location']) ? $_REQUEST['location'] : '';

    /* $UserID = 'IT-LOGON';
    $Password = 'SYSITSDADMIN'; */

    $data = new stdClass();
	$data->status = true;
	$data->data = "Data Test";

    $_sql = "SELECT USER_NAME, USER_ID  FROM SU_USER 
     WHERE UPPER(ACTIVE)='Y' 
     AND USER_ID='".$UserID."' 
     AND USER_PASSWORD = '".$Password."' 
     ORDER BY UPD_DATE DESC "; 

     $result_out = new stdClass(); //เข้าถึง class std
     $DataRows = []; //เป็น array
     $resultArray = array();
     if($Location == 'HATCH'){
        $connectionStr = $Location;
     }else{
        $connectionStr = 'RTS';
     }
     $DataRows = OracleSelectAll($connectionStr, $_sql);
     // มีเจอมากกว่า 1 row  
     if ($DataRows != false) {
         foreach ($DataRows as $row) {
             array_push($resultArray, $row);
         }
         $result_out->vdata = $resultArray;
         $result_out->status = (true);
         $result_out->status_login = (true);
         $result_out->message = 'Dupplicate Job !!';
         $result_out->error_code = '102';
         $result_out->error = '';
         wh_log_path(LogFolder, 'Dupplicate Job !!');
     }
     // ไม่เจอ  
     else {
         $resultArray = array();
         $result_out->vdata = $resultArray;
         $result_out->status = (false);
         $result_out->status_login = (false);
         $result_out->message = 'Job not found !!';
         $result_out->error_code = '101';
         $result_out->error = 'ES';
         wh_log_path(LogFolder, 'Job not found !!');
     }
     $LoginArray =  [
             "user" => $UserID,
             "password" => $Password,
             "location" => $Location,
             "connection" => $connectionStr
     ];
     $resultArray2 = array();
     /* $result_out->status = (true);
     $result_out->status_login = (true); */
     array_push($resultArray2, $LoginArray);
     $result_out->xdata = $resultArray2;
     $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

     return $response; // ส่งคำตอบกลับ
 });

$app->run(); // สั่งให้ระบบทำงาน

?>