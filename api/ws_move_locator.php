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

define("LogFolder", "LogPdaMoveloadtor");
// $LogFolder = "LogWSLoadingEntry";
// $LogExtensionFile = ".log";
// $LogFileNameFormat = "yyyyMMdd";
$RetPass = "0";

$app = new \Slim\App; // สร้าง object หลักของระบบ


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


$app->get('/', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $response->getBody()->write("Webservice PDA WCS - FNEntry (Sewing Entry)"); // สร้างคำตอบกลับ
    //wh_log("Webservice PDA WCS - FNEntry (Sewing Entry)");

    return $response; // ส่งคำตอบกลับ
});


$app->post('/SCAN_MOVE_BARCODE', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $locator = isset($_REQUEST['locator']) ? $_REQUEST['locator'] : '';
    $barcode = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';

    $someArray = [
        [
            "locator" => $locator,
            "barcode" => $barcode,
        ]
    ];

    $result_out = new stdClass();
    $DataRows = [];
    $op_error_code = '';
    $op_error_msge = '';

    $connectionStr = 'HATCH';
    $conn = GetConnectByLocation($connectionStr);

    if (!$conn) {
        $_err = oci_error();
        //trigger_error(htmlentities($_err['message'], ENT_QUOTES), E_USER_ERROR);
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "ES";
        $result_out->message = "Error Connect Database! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ
        wh_log_path(LogFolder, $_err['message']);

        return $response; // ส่งคำตอบกลับ

    } else {
        try {
            // Bind the your input and output parameters to PHP variables
            $command = 'BEGIN PDA_BARCODE_MOVE_LOC (:ip_user, :ip_loc, :ip_bar,:op_ErrorCode, :op_ErrorMsgs); END;';
            $stmt = oci_parse($conn, $command);

            // Bind the input parameter
            oci_bind_by_name($stmt, ":ip_user", $locator);
            oci_bind_by_name($stmt, ":ip_loc", $locator);
            oci_bind_by_name($stmt, ":ip_bar", $barcode);
            //oci_bind_by_name($stmt, ":WK_RECEIVED", $locator);
            //oci_bind_by_name($stmt, ":WK_USER", $barcode);

            // Bind the output parameter
            oci_bind_by_name($stmt, ":op_ErrorCode", $op_error_code, 200);
            oci_bind_by_name($stmt, ":op_ErrorMsgs", $op_error_msge, 1000);

            $_resExc = oci_execute($stmt);
            $_err = oci_error($stmt);
            $_resFree = oci_free_statement($stmt);
            if ($_resExc == false) {
                $op_error_code = "ES";
                $op_error_msge = "Error Execute Scan Check Barcode! : " . $_err['message'];
                wh_log_path(LogFolder, $op_error_msge);
                // return
                $result_out->status = (false);
            } else if ($op_error_code == $RetPass) {
                $DataRows = [
                    'error_code' => $op_error_code,
                    'error_msge' => $op_error_msge
                ];
                $_log = "pda_qc_output_check_barcode() USER=" . $locator . ",BARCODE=" . $barcode . "LOCATOR=" . $locator . "-" . "LOG=" . $op_error_code . "-" . $op_error_msge;
                wh_log_path(LogFolder, $op_error_msge);
                $result_out->status = (true);
            } else {
                $result_out->status = (true);
                $_log = "Error pda_qc_output_check_barcode() USER=" . $locator . ",BARCODE=" . $barcode . "-" . $op_error_code . "-" . $op_error_msge;
                wh_log_path(LogFolder, $_log);
            }
            //oci_close($conn);
        } catch (Throwable $e) {
            // Executed only in PHP 7, will not match in PHP 5.x
            $result_out->status = (false);
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Throwable Scan Check Barcode! " . $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } catch (Exception $e) {
            // Executed only in PHP 5.x, will not be reached in PHP 7
            $result_out->status = (false);
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Exception Scan Check Barcode! " . $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } finally {
            oci_close($conn);
        }
    }

    $DataRows_set = [
        'locator' => $locator,
        'barcode' => $barcode
    ];

    $resultArray = array();
    $resultArray2 = array();
    array_push($resultArray, $DataRows);
    array_push($resultArray2, $DataRows_set);
    //$result_out->status = (true);
    //$result_out->status = ($op_error_code == $RetPass);
    $result_out->error = $op_error_code;
    $result_out->message = $op_error_msge;
    $result_out->vdata = $resultArray;
    $result_out->l_data = $resultArray2;

    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->run();
