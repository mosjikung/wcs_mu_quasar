<?php
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: access");
header("Access-Control-Allow-Methods: GET,POST");
header("Access-Control-Allow-Credentials: true");
header('Content-Type: application/json;charset=utf-8');

use \Psr\Http\Message\ResponseInterface as Response; // ไลบราลี้สำหรับจัดการคำร้องขอ
use \Psr\Http\Message\ServerRequestInterface as Request;

// ไลบราลี้สำหรับจัดการคำตอบกลับ

require './vendor/autoload.php'; // ดึงไฟ์ autoload.php เข้ามา
//include_once './class.oracle.php'; // Class Connect Oracle
include_once './util.php'; // ดึงไฟ์ util.php เข้ามา
include_once './web_config.php';

define("LogFolder", "LogWSQCOutput");
// $LogFolder = "LogWSLoadingEntry";
// $LogExtensionFile = ".log";
// $LogFileNameFormat = "yyyyMMdd";
$RetPass = "0";

$app = new \Slim\App; // สร้าง object หลักของระบบ

$app->get('/', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $response->getBody()->write("Webservice PDA WCS - QCOutput"); // สร้างคำตอบกลับ
    //wh_log("Webservice PDA WCS - QCOutput");

    return $response; // ส่งคำตอบกลับ
});

//Start Connect Oracle
/**
 * @param  $_query
 * @return array      or false
 */

//End Connect Oracle

//Start Main App

// Check Role
function OracleSelectAll($_baselocation, $_query)
{
    $conn = GetConnectByLocation($_baselocation);
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
            $nrows = oci_fetch_all($stid, $DataRows, null, null, OCI_FETCHSTATEMENT_BY_ROW);
        } else {
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


// Check Role
$app->post('/checkrole', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $ulocation = 'HATCH';

    //$_sql = "SELECT * FROM DFIT_CONTROL_LOGISTIC";
    $locator = '381801800039';
    $_sql = @"SELECT H.PACKING_NO PL_NO,INV_NO,PO_NO,LINE_ID,SO_NO,CUSTOMER_ID, 
    CUSTOMER_DESC CUSTOMER_NAME,H.ITEM_CODE,H.COLOR_CODE COLOR_ID,H.ITEM_DESC, 
    DECODE(UOM,'KG','KGS','DOZ','DOZ',UOM) UOM, TOTAL_ROLL, 
    TOTAL_WEIGHT TOTAL_QTY,DH_PL_NO,DH_PL_OU 
    FROM NYF_FGPK_HEADER_V H
    WHERE H.PK_TYPE='01.ส่งผ้าสีออกให้ลูกค้า' 
        AND (SUBSTR(H.PACKING_NO,2,1) = '1' OR SUBSTR(H.PACKING_NO,2,1)='8') 
        AND NVL(H.CANCEL_ACTIVE,'N')='N'  
        AND PACKING_NO='".$locator."' 
        AND H.ENTRY_DATE >= TO_DATE('01/01/2018','DD/MM/RRRR')  
        AND GET_REMAIN_PL (PACKING_NO,PO_NO)='PASS' ";

    $result_out = new stdClass();

    $DataRows = [];
    $resultArray = array();
    $DataRows = OracleSelectAll($ulocation, $_sql);
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            //echo $row['VERSION'] . '<br>';
            array_push($resultArray, $row);
        }
        
        $result_out->status = true;
        $result_out->error = "";
        $result_out->message = "";
        $result_out->vdata = $resultArray;
    } else {
        $resultArray = array();
        $result_out->status = true;
        $result_out->error = "ES";
        $result_out->message = "USERNAME NOT ACCESS IS DENIED. PLEASE CONTACT IT";
        $result_out->vdata = $resultArray;
    }
    $response->getBody()->write(json_encode($result_out));

    return $response; // ส่งคำตอบกลับ
});



$app->run(); // สั่งให้ระบบทำงาน
