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

define("LogFolder", "LogWSmovelocator");
// $LogFolder = "LogWSLoadingEntry";
// $LogExtensionFile = ".log";
// $LogFileNameFormat = "yyyyMMdd";
$RetPass = "0";

$app = new \Slim\App; // สร้าง object หลักของระบบ

$app->get('/', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $response->getBody()->write("Webservice PDA WCS - FNEntry (Sewing Entry)"); // สร้างคำตอบกลับ
    //wh_log("Webservice PDA WCS - FNEntry (Sewing Entry)");

    return $response; // ส่งคำตอบกลับ
});

// check connect and query
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

function GetItemCodeByBarcode($WKBARCODE, $connectionStr)
{

    $WKITEM = '';
    $_sql = @"SELECT PO_ITEM_CODE FROM DFWH_WAREHOUSE WHERE BARCODE = '" . $WKBARCODE . "' AND NVL(FABRIC_LOCK,'N')='N' 
    ";
    $DataRows = [];
    $DataRows = OracleSelectAll($connectionStr, $_sql);
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            $WKITEM = $row['PO_ITEM_CODE'];
        }
    }

    return $WKITEM;
}

function GetUSER_GROUP_LOC($P_USE, $connectionStr)
{

    $vWH_TYPE = '';
    $_sql = @"SELECT  EMP_ID , DECODE(EMP_GROUP,'OMNOI','TC OMNOI','TC WATSON')  WH_TYPE FROM DFIT_USER  WHERE EMP_ID = '" . $P_USE . "' 
    ";
    $DataRows = [];
    $DataRows = OracleSelectAll($connectionStr, $_sql);
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            $vWH_TYPE = $row['WH_TYPE'];
        }
    }

    return $vWH_TYPE;
}

function SCAN_CHK_JOB($P_JOB, $vWH_TYPE, $connectionStr)
{
    try {
        $check_status = '';
        $_sql = @"SELECT H.JOB_NO,H.PACKING_NO,H.JOB_COMPLETE FROM WEB_JOB_PDA H 
    WHERE  H.JOB_NO='" . $P_JOB . "' AND H.WH_TYPE ='" . $vWH_TYPE . "' 
    ";
        if ($vWH_TYPE == '') {
            $_sql += " AND 1=0";
        }
        $DataRows = [];
        $DataRows = OracleSelectAll($connectionStr, $_sql);
        if ($DataRows != false) {
            return $DataRows;
        } else {
            $DataRows = false;
        }
        return $check_status;
    } catch (Exception $e) {
        return $DataRows = false;
    }
}

function Get_Count_LOC($P_JOB, $connectionStr)
{
    $count = '';
    $_sql = @"SELECT  REQ_JOB_NO, C_COUNT FROM RTS_REQ_TRANSFER_DV WHERE REQ_JOB_NO = '".$P_JOB."' ";
    $DataRows = [];
    $DataRows = OracleSelectAll($connectionStr, $_sql);
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            $count = $row['C_COUNT'];
        }
    }
    return $count;
}

// Check Role
$app->post('/PDA_LOC', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $CON_STR = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : '';
    $connectionStr = $CON_STR;

    $_sql = "SELECT 'GN1F1-01' as LOC_NO , 'GN1F1-01' as LOC_DESC FROM DUAL";

    $DataRows = OracleSelectAll($connectionStr, $_sql);
    $resultArray = array();
    $count = 0;
    foreach ($DataRows as $row) {
        /* $SetLocation =  [
            "LOCATION" => $row,
            "values" => $count
        ];
        $_location = new stdClass(); */
        array_push($resultArray, $row);
        $count++;
    }
    $result_out = new stdClass();
    $result_out->status = true;
    $result_out->error = "";
    $result_out->message = "";
    $result_out->vdata = $resultArray;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response;
});
$app->post('/GET_BARCODE_SCAN', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $P_JOB = isset($_REQUEST['job_no']) ? $_REQUEST['job_no'] : '';
    $CON_STR = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : '';
    $P_USE = isset($_REQUEST['Username']) ? $_REQUEST['Username'] : '';
    //$P_JOB = 'RQ21090014';
    $vWH_TYPE = '';


    $_sql = "SELECT BARCODE_ID as BARCODE, ITEM_CODE, NVL(UNIT_QTY,0) as QTY FROM RTS_REQ_TRANSFER_D 
    WHERE REQ_JOB_NO= '" . $P_JOB . "' AND NVL(CONFIRM_REQ,'N')='Y' AND BARCODE_ID IS NOT NULL ORDER BY ENTRY_DATE DESC";

    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array();
    $connectionStr = $CON_STR;
    $Count = Get_Count_LOC($P_JOB, $connectionStr);
    $DataRows = OracleSelectAll($connectionStr, $_sql);
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            array_push($resultArray, $row);
        }
        $result_out->vdata = $resultArray;
        $result_out->status = (true);
        $result_out->Count = $Count;
        $result_out->vWH_TYPE = $vWH_TYPE;
        $result_out->message = 'Job true !!';
        $result_out->error_code = '102';
        $result_out->error = '';
    } else {
        $result_out->vdata = $resultArray;
        $result_out->status = (false);
        $result_out->Count = $Count;
        $result_out->vWH_TYPE = $vWH_TYPE;
        $result_out->message = 'fail Job !!';
        $result_out->error_code = '102';
        $result_out->error = '';
    }
    /*  $someArray = [
        [
            "job_no" => $P_JOB,
            "connection" => $CON_STR,
            "username" => $P_USE
        ]
    ]; */
    /* $resultArray2 = array();
    array_push($resultArray2, $someArray);
    $result_out->P_JOB = $P_JOB;
    $result_out->l_data = $resultArray2; */
    $response->getBody()->write(json_encode($result_out));

    return $response; // ส่งคำตอบกลับ
});
$app->post('/GET_BARCODE_BALANCE', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $P_JOB = isset($_REQUEST['job_no']) ? $_REQUEST['job_no'] : '';
    $CON_STR = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : '';
    $P_USE = isset($_REQUEST['Username']) ? $_REQUEST['Username'] : '';
    //$P_JOB = 'JOB19030016';
    $vWH_TYPE = '';

    $_sql = "SELECT  H.JOB_NO,D.BARCODE, D.ITEM_CODE,D.STK_ITEM_CODE as ITEM_CODE_NYK , D.COLOR_CODE as COLOR_CODE_NYK ,
    D.STK_ITEM_CODE as ITEM_CODE_NYK , D.COLOR_CODE as COLOR_CODE_NYK ,
    D.FG_UNIT_QTY as QTY, 'Y' INV_CORRECT
    FROM NYF_FGPK_HEADER H, NYF_FGPK_DETAIL D WHERE H.PACKING_NO  = D.PACKING_NO AND H.JOB_NO = '" . $P_JOB . "'
    AND D.BARCODE NOT IN (SELECT S.BARCODE FROM PDA_SHIPMENT S WHERE  S.INV_CORRECT ='Y'
    AND S.JOB_NO = H.JOB_NO AND S.BARCODE = D.BARCODE ) ORDER BY 1 ";

    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array();
    $connectionStr = $CON_STR;

    $DataRows = OracleSelectAll($connectionStr, $_sql);
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            array_push($resultArray, $row);
        }
        $result_out->vdata = $resultArray;
        $result_out->status = (true);
        $result_out->vWH_TYPE = $vWH_TYPE;
        $result_out->message = 'Job true !!';
        $result_out->error_code = '102';
        $result_out->error = '';
    } else {
        $result_out->vdata = $resultArray;
        $result_out->status = (false);
        $result_out->vWH_TYPE = $vWH_TYPE;
        $result_out->message = 'fail Job !!';
        $result_out->error_code = '102';
        $result_out->error = '';
    }
    /*  $someArray = [
        [
            "job_no" => $P_JOB,
            "connection" => $CON_STR,
            "username" => $P_USE
        ]
    ]; */
    /* $resultArray2 = array();
    array_push($resultArray2, $someArray);
    $result_out->P_JOB = $P_JOB;
    $result_out->l_data = $resultArray2; */
    $response->getBody()->write(json_encode($result_out));

    return $response; // ส่งคำตอบกลับ
});
$app->post('/PDA_GET_LOC_USER', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $P_JOB = isset($_REQUEST['job_no']) ? $_REQUEST['job_no'] : '';
    $CON_STR = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : '';
    $P_USE = isset($_REQUEST['Username']) ? $_REQUEST['Username'] : '';
    //$P_JOB = 'JOB20030034';
    $vWH_TYPE = '';

    $_sql = "SELECT 'GN1F1-01' as LOC_NO , 'GN1F1-01' as LOC_DESC FROM DUAL";

    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array();
    $connectionStr = $CON_STR;

    $DataRows = OracleSelectAll($connectionStr, $_sql);
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            array_push($resultArray, $row);
        }
        $result_out->vdata = $resultArray;
        $result_out->status = (true);
        $result_out->vWH_TYPE = $vWH_TYPE;
        $result_out->message = 'Job true !!';
        $result_out->error_code = '102';
        $result_out->error = '';
    } else {
        $result_out->vdata = $resultArray;
        $result_out->status = (false);
        $result_out->vWH_TYPE = $vWH_TYPE;
        $result_out->message = 'fail Job !!';
        $result_out->error_code = '102';
        $result_out->error = '';
    }
    /*  $someArray = [
        [
            "job_no" => $P_JOB,
            "connection" => $CON_STR,
            "username" => $P_USE
        ]
    ]; */
    /* $resultArray2 = array();
    array_push($resultArray2, $someArray);
    $result_out->P_JOB = $P_JOB;
    $result_out->l_data = $resultArray2; */
    $response->getBody()->write(json_encode($result_out));

    return $response; // ส่งคำตอบกลับ
});
$app->post('/SCAN_JOB', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $P_JOB = isset($_REQUEST['job_no']) ? $_REQUEST['job_no'] : '';
    $CON_STR = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : '';
    $P_USE = isset($_REQUEST['Username']) ? $_REQUEST['Username'] : '';
    //$P_JOB = 'JOB20030034';
    $vWH_TYPE = '';

    $_sql = "SELECT JOB_NO,BARCODE,ITEM_CODE,FG_UNIT_QTY as QTY, DECODE(INV_CORRECT,'Y','True','False')    INV_CORRECT FROM PDA_SHIPMENT 
    WHERE JOB_NO = '" . $P_JOB . "' AND INV_CORRECT ='Y' ORDER BY ENTRY_DATE DESC ";

    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array();
    $connectionStr = $CON_STR;

    $DataRows = OracleSelectAll($connectionStr, $_sql);
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            array_push($resultArray, $row);
        }
        $result_out->vdata = $resultArray;
        $result_out->status = (true);
        $result_out->vWH_TYPE = $vWH_TYPE;
        $result_out->message = 'Job true !!';
        $result_out->error_code = '102';
        $result_out->error = '';
    } else {
        $result_out->vdata = $resultArray;
        $result_out->status = (false);
        $result_out->vWH_TYPE = $vWH_TYPE;
        $result_out->message = 'fail Job !!';
        $result_out->error_code = '102';
        $result_out->error = '';
    }
    /*  $someArray = [
        [
            "job_no" => $P_JOB,
            "connection" => $CON_STR,
            "username" => $P_USE
        ]
    ]; */
    /* $resultArray2 = array();
    array_push($resultArray2, $someArray);
    $result_out->P_JOB = $P_JOB;
    $result_out->l_data = $resultArray2; */
    $response->getBody()->write(json_encode($result_out));

    return $response; // ส่งคำตอบกลับ
});
$app->post('/SCAN_BARCODE', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $WKJOB = isset($_REQUEST['job_no']) ? $_REQUEST['job_no'] : '';
    $CON_STR = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : '';
    $WKBAR = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';
    $WKUSER = isset($_REQUEST['Username']) ? $_REQUEST['Username'] : '';
    //$WKJOB = 'JOB21060038';
    $op_ErrorCode = "";
    $op_ErrorMsgs = "";

    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array();
    $connectionStr = $CON_STR;
    $conn = GetConnectByLocation($connectionStr);
    $result_out->status = false;

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
            $command = '';
            // Bind the your input and output parameters to PHP variables
            $command = 'BEGIN PDA_CHK_BAR_REQ_JOB (
            :WKJOB,
            :WKBAR, 
            :WKUSER,
            :op_ErrorCode,
            :op_ErrorMsgs,
                                                        
            ); END;'; 
            $stmt = oci_parse($conn, $command);

            // Bind the input parameter
            oci_bind_by_name($stmt, ":WKJOB", $WKJOB);
            oci_bind_by_name($stmt, ":WKBAR", $WKBAR);
            oci_bind_by_name($stmt, ":WKUSER", $WKUSER);
            // Bind the output parameter
            oci_bind_by_name($stmt, ":op_ErrorCode", $op_ErrorCode, 200);
            oci_bind_by_name($stmt, ":op_ErrorMsgs", $op_ErrorMsgs, 200);

            $_resExc = oci_execute($stmt);
            $_err = oci_error($stmt);
            $_resFree = oci_free_statement($stmt);
            if ($_resExc == false) {
                $result_out->status = false;
                $op_ErrorCode = "ES";
                $op_ErrorMsgs = "Error Execute Cancel_barcode ! : " . $_err['message'];
                wh_log_path(LogFolder, $op_ErrorMsgs);
                // return
            } else if ($op_ErrorCode == $RetPass) {
                $result_out->status = true;
                $DataRows = [
                    'error_code' => $op_ErrorCode,
                    'error_msge' => $op_ErrorMsgs
                ];
                array_push($resultArray, $DataRows);
                $_log = "Cancel_barcode log // WKJOB=" . $WKJOB . ",WKUSER=" . $WKUSER . "WKBARCODE=" . $WKBAR . "-" . "LOG=" . $op_ErrorCode . "-" . $op_ErrorMsgs;
                wh_log_path(LogFolder, $op_ErrorMsgs);
            } else {
                $result_out->status = false;
                $_log = "Error Cancel_barcode log // WKJOB=" . $WKJOB . ",WKUSER=" . $WKUSER . "WKBARCODE=" . $WKBAR . "-" . "LOG=" . $op_ErrorCode . "-" . $op_ErrorMsgs;
                wh_log_path(LogFolder, $_log);
            }

            //oci_close($conn);
        } catch (Throwable $e) {
            // Executed only in PHP 7, will not match in PHP 5.x
            $result_out->status = false;
            $_res = $e;
            $op_ErrorCode = "ES";
            $op_ErrorMsgs = "Error Throwable Scan Check Barcode! " . $_res;
            wh_log_path(LogFolder, $op_ErrorMsgs);
        } catch (Exception $e) {
            // Executed only in PHP 5.x, will not be reached in PHP 7
            $result_out->status = false;
            $_res = $e;
            $op_ErrorCode = "ES";
            $op_ErrorMsgs = "Error Exception Scan Check Barcode! " . $_res;
            wh_log_path(LogFolder, $op_ErrorMsgs);
        } finally {
            oci_close($conn);
        }
    }

    $result_out->status = false;
    $result_out->error = $op_ErrorCode;
    $result_out->message = $op_ErrorMsgs;
    $result_out->vdata = $resultArray;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});
$app->post('/CONFIRM_JOB', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $WKJOB = isset($_REQUEST['job_no']) ? $_REQUEST['job_no'] : '';
    $CON_STR = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : '';
    $WKBARCODE = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';
    $WKUSER = isset($_REQUEST['Username']) ? $_REQUEST['Username'] : '';
    //$WKJOB = 'JOB21060038';


    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array();
    $connectionStr = $CON_STR;
    $conn = GetConnectByLocation($connectionStr);
    $result_out->status = false;


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
            $command = '';
            // Bind the your input and output parameters to PHP variables
            $command = 'BEGIN PDA_CONF_REQ_JOB (:WKJOB, :WKUSER,:op_ErrorCode, :op_ErrorMsgs); END;';
            $stmt = oci_parse($conn, $command);
            // Bind the input parameter
            oci_bind_by_name($stmt, ":WKJOB", $WKJOB);
            oci_bind_by_name($stmt, ":WKUSER", $WKUSER);
            // Bind the output parameter
            oci_bind_by_name($stmt, ":op_ErrorCode", $op_ErrorCode, 200);
            oci_bind_by_name($stmt, ":op_ErrorMsgs", $op_ErrorMsgs, 200);
            $_resExc = oci_execute($stmt);
            $_err = oci_error($stmt);
            $_resFree = oci_free_statement($stmt);
            if ($_resExc == false) {
                $result_out->status = false;
                $op_ErrorCode = "ES";
                $op_ErrorMsgs = "Error Execute Cancel_barcode ! : " . $_err['message'];
                wh_log_path(LogFolder, $op_ErrorMsgs);
                // return
            } else if ($op_ErrorCode == $RetPass) {
                $DataRows = [
                    'error_code' => $op_ErrorCode,
                    'error_msge' => $op_ErrorMsgs
                ];
                array_push($resultArray, $DataRows);
                $result_out->status = true;
                $_log = "Cancel_barcode log // WKJOB=" . $WKJOB . ",WKUSER=" . $WKUSER . "WKBARCODE=" . $WKBARCODE . "-" . "LOG=" . $op_ErrorCode . "-" . $op_ErrorMsgs;
                wh_log_path(LogFolder, $op_ErrorMsgs);
            } else {
                $result_out->status = false;
                $_log = "Error Cancel_barcode log // WKJOB=" . $WKJOB . ",WKUSER=" . $WKUSER . "WKBARCODE=" . $WKBARCODE . "-" . "LOG=" . $op_ErrorCode . "-" . $op_ErrorMsgs;
                wh_log_path(LogFolder, $_log);
            }
            //oci_close($conn);
        } catch (Throwable $e) {
            // Executed only in PHP 7, will not match in PHP 5.x
            $result_out->status = false;
            $_res = $e;
            $op_ErrorCode = "ES";
            $op_ErrorMsgs = "Error Throwable Scan Check Barcode! " . $_res;
            wh_log_path(LogFolder, $op_ErrorMsgs);
        } catch (Exception $e) {
            // Executed only in PHP 5.x, will not be reached in PHP 7
            $result_out->status = false;
            $_res = $e;
            $op_ErrorCode = "ES";
            $op_ErrorMsgs = "Error Exception Scan Check Barcode! " . $_res;
            wh_log_path(LogFolder, $op_ErrorMsgs);
        } finally {
            oci_close($conn);
        }
    }


    $result_out->error = $op_ErrorCode;
    $result_out->message = $op_ErrorMsgs;
    $result_out->vdata = $resultArray;


    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->run();
