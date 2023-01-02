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

define("LogFolder", "LogWSLoadingEntry");
// $LogFolder = "LogWSLoadingEntry";
// $LogExtensionFile = ".log";
// $LogFileNameFormat = "yyyyMMdd";
$RetPass = "0";

$app = new \Slim\App; // สร้าง object หลักของระบบ

$app->get('/', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $response->getBody()->write("Webservice PDA WCS - LoadingEntry"); // สร้างคำตอบกลับ
    //wh_log("Webservice PDA WCS - LoadingEntry");

    return $response; // ส่งคำตอบกลับ
});

//Start Connect Oracle

/**
 * @param  $_query
 * @return array      or false
 */
function OracleSelectAll($_baselocation, $_query)
{
    $conn = GetConnectByLocation($_baselocation);
    $DataRows = array();

    if (!$conn) {
        $_err = oci_error();
        $_res = $_err['message'];
        wh_log_path(LogFolder, $_res);
        // $result_out = new stdClass();
        // $resultArray = array();
        // $result_out->status = true;
        // $result_out->error = "ES ERROR CONNECT DATABASE";
        // $result_out->message = $m['message'];
        // $result_out->vdata = $resultArray;
        // return $m['message'];
        //echo $m['message'], "\n";
        //exit;
    }

    try {

        $stid = oci_parse($conn, $_query);
        $_res = oci_execute($stid);
        if ($_res) {
            $nrows = oci_fetch_all($stid, $DataRows, null, null, OCI_FETCHSTATEMENT_BY_ROW);
        } else {
            $DataRows = false;
            // $DataRows = [];
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

//End Connect Oracle

//Start Main App

// Check Role
$app->post('/checkrole', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';

    $_sql = "   SELECT EMP_NAME USER_NAME,EMP_CODE USER_ID
                FROM DB_USER
                WHERE UPPER(ACTIVE)='Y'
                    AND LOAD ='Y'
                    AND EMP_CODE='$uid'
                ORDER BY  UPD_DATE DESC ";

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
        $result_out->status = false;
        $result_out->error = "ES";
        $result_out->message = "USERNAME NOT ACCESS IS DENIED. PLEASE CONTACT IT";
        $result_out->vdata = $resultArray;
    }
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/scanbarcode', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';

    $so_no = isset($_REQUEST['so_no']) ? $_REQUEST['so_no'] : '';
    $so_year = isset($_REQUEST['so_year']) ? $_REQUEST['so_year'] : '';
    $barcode = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';

    $result_out = new stdClass();
    $DataRows = [];
    $op_error_code = '';
    $op_error_msge = '';

    if ($barcode == '') {

        $strlog = "ENTRY BARCODE ! USER=" . $uid . ",BARCODE=" . $barcode;
        //wh_log_path(LogFolder, $strlog);

        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "PLEASE INSERT BARCODE ! ";
        $result_out->message = "PLEASE INSERT BARCODE ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        return $response; // ส่งคำตอบกลับ
    }

    // Connect to database
    $conn = GetConnectByLocation($ulocation);
    // Through error if not connected
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
            $command = 'BEGIN pda_loaing_entry_check_barcode(
                                :ip_empuser, :ip_so_no, :ip_so_year, :ip_barcode,
                                :op_error_code, :op_error_msge,
                                :op_barcode, :op_so_no, :op_so_year, :op_customer_id, :op_customer_name,
                                :op_plan_id, :op_style, :op_line, :op_color, :op_size, :op_qty
                    ); END;';
            $stmt = oci_parse($conn, $command);

            // Bind the input parameter
            oci_bind_by_name($stmt, ":ip_empuser", $uid);
            oci_bind_by_name($stmt, ":ip_so_no", $so_no);
            oci_bind_by_name($stmt, ":ip_so_year", $so_year);
            oci_bind_by_name($stmt, ":ip_barcode", $barcode);

            // Bind the output parameter
            oci_bind_by_name($stmt, ":op_error_code", $op_error_code, 100);
            oci_bind_by_name($stmt, ":op_error_msge", $op_error_msge, 200);

            oci_bind_by_name($stmt, ":op_barcode", $op_barcode, 200);
            oci_bind_by_name($stmt, ":op_so_no", $op_so_no, 40);
            oci_bind_by_name($stmt, ":op_so_year", $op_so_year, 40);
            oci_bind_by_name($stmt, ":op_customer_id", $op_customer_id, 200);
            oci_bind_by_name($stmt, ":op_customer_name", $op_customer_name, 200);
            oci_bind_by_name($stmt, ":op_plan_id", $op_plan_id, 50);
            oci_bind_by_name($stmt, ":op_style", $op_style, 50);
            oci_bind_by_name($stmt, ":op_line", $op_line, 50);
            oci_bind_by_name($stmt, ":op_color", $op_color, 50);
            oci_bind_by_name($stmt, ":op_size", $op_size, 40);
            oci_bind_by_name($stmt, ":op_qty", $op_qty, 40);

            $_resExc = oci_execute($stmt);
            $_err = oci_error($stmt);
            $_resFree = oci_free_statement($stmt);
            if ($_resExc == false) {
                $op_error_code = "ES";
                $op_error_msge = "Error Execute Scan Check Barcode! : " . $_err['message'];
                //$op_error_msge = "Error Execute Scan Check Barcode! : ". $_err['code'] ." , ". $_err['message']." , ". $_err['offset']+1." , ". $_err['sqltext'];
                wh_log_path(LogFolder, $op_error_msge);
                // return
            } else if ($op_error_code == $RetPass) {
                $DataRows = [
                    'error_code' => $op_error_code,
                    'error_msge' => $op_error_msge,

                    'barcode' => $op_barcode,
                    'so_year' => $op_so_year,
                    'so_no' => $op_so_no,
                    'customer_id' => $op_customer_id,
                    'customer_name' => $op_customer_name,
                    'plan_id' => $op_plan_id,
                    'style' => $op_style,
                    'line' => $op_line,
                    'color' => $op_color,
                    'size' => $op_size,
                    'qty' => $op_qty,
                ];
            } else {
                $_log = "Error pda_loaing_entry_check_barcode() USER=" . $uid . ",BARCODE=" . $barcode . " - " . $op_error_code . "-" . $op_error_msge;
                wh_log_path(LogFolder, $_log);
            }
            //oci_close($conn);
        } catch (Throwable $e) {
            // Executed only in PHP 7, will not match in PHP 5.x
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Throwable Scan Check Barcode! " . $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } catch (Exception $e) {
            // Executed only in PHP 5.x, will not be reached in PHP 7
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Exception Scan Check Barcode! " . $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } finally {
            oci_close($conn);
        }
    }

    $resultArray = array();
    array_push($resultArray, $DataRows);
    $result_out->status = ($op_error_code == $RetPass);
    $result_out->error = $op_error_code;
    $result_out->message = $op_error_msge;
    $result_out->vdata = $resultArray;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/scanbarcodeconfirm', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';

    $so_no = isset($_REQUEST['so_no']) ? $_REQUEST['so_no'] : '';
    $so_year = isset($_REQUEST['so_year']) ? $_REQUEST['so_year'] : '';
    $barcode = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';

    $result_out = new stdClass();
    $DataRows = [];
    $op_error_code = '';
    $op_error_msge = '';

    if ($barcode == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "ES";
        $result_out->message = "PLEASE INSERT BARCODE ! ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY BARCODE ! USER=" . $uid . ",BARCODE=" . $barcode;
        //wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($so_year == '' || $so_no == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "ES";
        $result_out->message = "Error Entry SO No. or  SO Year ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY SO ! USER=" . $uid . ",BARCODE=" . $barcode;
        //wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    // Connect to database
    $conn = GetConnectByLocation($ulocation);
    // Through error if not connected
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
            $command = 'BEGIN pda_loaing_entry_cfm_barcode(
                                :ip_empuser, :ip_so_no, :ip_so_year, :ip_barcode,
                                :op_error_code, :op_error_msge,
                                :op_barcode, :op_so_no, :op_so_year, :op_customer_id, :op_customer_name,
                                :op_plan_id, :op_style, :op_line, :op_color, :op_size, :op_qty
                    ); END;';
            // $command = 'BEGIN pda_loaing_entry_check_barcode(
            //     :ip_empuser, :ip_so_no, :ip_so_year, :ip_barcode,
            //     :op_error_code, :op_error_msge,
            //     :op_barcode, :op_so_no, :op_so_year, :op_customer_id, :op_customer_name,
            //     :op_plan_id, :op_style, :op_line, :op_color, :op_size, :op_qty
            // ); END;';
            $stmt = oci_parse($conn, $command);

            // Bind the input parameter
            oci_bind_by_name($stmt, ":ip_empuser", $uid);
            oci_bind_by_name($stmt, ":ip_so_no", $so_no);
            oci_bind_by_name($stmt, ":ip_so_year", $so_year);
            oci_bind_by_name($stmt, ":ip_barcode", $barcode);

            // Bind the output parameter
            oci_bind_by_name($stmt, ":op_error_code", $op_error_code, 100);
            oci_bind_by_name($stmt, ":op_error_msge", $op_error_msge, 200);

            oci_bind_by_name($stmt, ":op_barcode", $op_barcode, 200);
            oci_bind_by_name($stmt, ":op_so_no", $op_so_no, 40);
            oci_bind_by_name($stmt, ":op_so_year", $op_so_year, 40);
            oci_bind_by_name($stmt, ":op_customer_id", $op_customer_id, 200);
            oci_bind_by_name($stmt, ":op_customer_name", $op_customer_name, 200);
            oci_bind_by_name($stmt, ":op_plan_id", $op_plan_id, 50);
            oci_bind_by_name($stmt, ":op_style", $op_style, 50);
            oci_bind_by_name($stmt, ":op_line", $op_line, 50);
            oci_bind_by_name($stmt, ":op_color", $op_color, 50);
            oci_bind_by_name($stmt, ":op_size", $op_size, 40);
            oci_bind_by_name($stmt, ":op_qty", $op_qty, 40);

            $_resExc = oci_execute($stmt);
            $_err = oci_error($stmt);
            $_resFree = oci_free_statement($stmt);
            if ($_resExc == false) {
                $op_error_code = "ES";
                $op_error_msge = "Error Execute Scan Confirm Barcode! : " . $_err['message'];
                wh_log_path(LogFolder, $op_error_msge);
            } else if ($op_error_code == $RetPass) {
                $_log = "BARCODE_CONFIRM " . $barcode;
                wh_log_path(LogFolder, $_log);

                $DataRows = [
                    'error_code' => $op_error_code,
                    'error_msge' => $op_error_msge,

                    'barcode' => $op_barcode,
                    'so_year' => $op_so_year,
                    'so_no' => $op_so_no,
                    'customer_id' => $op_customer_id,
                    'customer_name' => $op_customer_name,
                    'plan_id' => $op_plan_id,
                    'style' => $op_style,
                    'line' => $op_line,
                    'color' => $op_color,
                    'size' => $op_size,
                    'qty' => $op_qty,
                ];
            } else {
                $_log = "Error pda_loaing_entry_cfm_barcode() USER=" . $uid . ",BARCODE=" . $barcode . " - " . $op_error_code . "-" . $op_error_msge;
                wh_log_path(LogFolder, $_log);
            }
            //oci_close($conn);
        } catch (Throwable $e) {
            // Executed only in PHP 7, will not match in PHP 5.x
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Throwable Scan Confirm Barcode! " . $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } catch (Exception $e) {
            // Executed only in PHP 5.x, will not be reached in PHP 7
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Exception Scan Confirm Barcode! " . $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } finally {
            oci_close($conn);
        }
    }

    $resultArray = array();
    array_push($resultArray, $DataRows);
    $result_out->status = ($op_error_code == $RetPass);
    $result_out->error = $op_error_code;
    $result_out->message = $op_error_msge;
    $result_out->vdata = $resultArray;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->run(); // สั่งให้ระบบทำงาน
