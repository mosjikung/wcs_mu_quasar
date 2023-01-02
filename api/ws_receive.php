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

define("LogFolder", "LogPdaReceive");
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
        // $_res = true;
        $e = oci_error($stid);
        if ($_res) {
            $nrows = oci_fetch_all($stid, $DataRows, null, null, OCI_FETCHSTATEMENT_BY_ROW);

            wh_log_path(LogFolder, 'TrueX = ' . $e);
        } else {
            $DataRows = false;
            wh_log_path(LogFolder, 'FalseX = ' . $e);
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

function GET_SUMMARY_PACKING($_baselocation, $PACKING_NO)
{
    $conn = GetConnectByLocation($_baselocation);
    $DataRows = array();
    $_query = "SELECT COUNT(*) CNT,NVL(SUM(UNIT_QTY),0) KG FROM DFWH_WAREHOUSE WHERE PACKING_NO='" . $PACKING_NO . "' ";

    if (!$conn) {
        $_err = oci_error();
        $_res = $_err['message'];
        wh_log_path(LogFolder, $_res);
    }
    try {
        $stid = oci_parse($conn, $_query);
        $_res = oci_execute($stid);
        // $_res = true;
        $e = oci_error($stid);
        if ($_res) {
            $nrows = oci_fetch_all($stid, $DataRows, null, null, OCI_FETCHSTATEMENT_BY_ROW);
            wh_log_path(LogFolder, 'True = ' . $e);
        } else {
            $DataRows = false;
            wh_log_path(LogFolder, 'False = ' . $e);
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
$app->post('/check_packing', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $packing = isset($_REQUEST['packing']) ? $_REQUEST['packing'] : '';
    $connectionStr = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : 'HATCH';

    //$packing = '381801800039';
    $_sql = @"SELECT H.PACKING_NO PL_NO,INV_NO,PO_NO,LINE_ID,SO_NO,CUSTOMER_ID, 
    CUSTOMER_DESC CUSTOMER_NAME,H.ITEM_CODE,H.COLOR_CODE COLOR_ID,H.ITEM_DESC, 
    DECODE(UOM,'KG','KGS','DOZ','DOZ',UOM) UOM, TOTAL_ROLL, 
    TOTAL_WEIGHT TOTAL_QTY,DH_PL_NO,DH_PL_OU 
    FROM NYF_FGPK_HEADER_V H
    WHERE H.PK_TYPE='01.ส่งผ้าสีออกให้ลูกค้า' 
        AND (SUBSTR(H.PACKING_NO,2,1) = '1' OR SUBSTR(H.PACKING_NO,2,1)='8') 
        AND NVL(H.CANCEL_ACTIVE,'N')='N'  
        AND PACKING_NO='" . $packing . "' 
        AND H.ENTRY_DATE >= TO_DATE('01/01/2018','DD/MM/RRRR')  
        AND GET_REMAIN_PL (PACKING_NO,PO_NO)='PASS' ";

    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array();
    //$connectionStr = 'HATCH';
    $DataRows = OracleSelectAll($connectionStr, $_sql);
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            array_push($resultArray, $row);
        }
        $result_out->vdata = $resultArray;
        $result_out->status = (true);
        $result_out->message = 'Dupplicate Job !!';
        $result_out->error_code = '102';
        $result_out->packing = $packing;
        $result_out->connectionStr = $connectionStr;
        $result_out->error = '';
        wh_log_path(LogFolder, 'true Job !!');
    } else {
        //$resultArray = array();
        //$result_out->vdata = $resultArray;
        $result_out->vdata = [];
        $result_out->status = (false);
        $result_out->message = 'Job not found !!';
        $result_out->error_code = '101';
        $result_out->packing = $packing;
        $result_out->connectionStr = $connectionStr;
        $result_out->error = 'ES';
        wh_log_path(LogFolder, 'fail Job !!');
    }

    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/scan_barcode', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $WK_LOCATER = isset($_REQUEST['locator']) ? $_REQUEST['locator'] : '';
    $WK_PL = isset($_REQUEST['packing']) ? $_REQUEST['packing'] : '';
    $WK_BARCODE = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';
    $WK_RECEIVED = isset($_REQUEST['receive']) ? $_REQUEST['receive'] : '';
    $WKUSER = isset($_REQUEST['Username']) ? $_REQUEST['Username'] : '';
    $connectionStr = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : 'HATCH';

    $result_out = new stdClass();
    $DataRows = [];
    $op_Received = '';
    $op_ErrorCode = '';
    $op_ErrorMsgs = '';
    $WK_PO = '';
    $WK_LINE = '';
    $connectionStr = 'HATCH';
    $conn = GetConnectByLocation($connectionStr);

    $open = false;
    if ($open) {
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
                $command = 'BEGIN PDA_SCAN_RECV (
            :WK_PO,
            :WK_LINE, 
            :WK_LOCATER,
            :WK_PL,
            :WK_BARCODE,
            :WK_RECEIVED,
            :WKUSER,
            :op_Received,
            :op_ErrorCode,
            :op_ErrorMsgs
            ); END;';
                //$command = 'BEGIN PDA_SCAN_RECV_C (:WK_RECEIVED, :WK_USER,:op_ErrorCode, :op_ErrorMsgs); END;';
                $stmt = oci_parse($conn, $command);

                // Bind the input parameter
                oci_bind_by_name($stmt, ":WK_PO", $WK_PO);
                oci_bind_by_name($stmt, ":WK_LINE", $WK_LINE);
                oci_bind_by_name($stmt, ":WK_LOCATER", $WK_LOCATER);
                oci_bind_by_name($stmt, ":WK_PL", $WK_PL);
                oci_bind_by_name($stmt, ":WK_BARCODE", $WK_BARCODE);
                oci_bind_by_name($stmt, ":WK_RECEIVED", $WK_RECEIVED);
                oci_bind_by_name($stmt, ":WKUSER", $WKUSER);
                //oci_bind_by_name($stmt, ":WK_RECEIVED", $locator);
                //oci_bind_by_name($stmt, ":WK_USER", $barcode);

                // Bind the output parameter
                oci_bind_by_name($stmt, ":op_Received", $op_Received, 200);
                oci_bind_by_name($stmt, ":op_ErrorCode", $op_ErrorCode, 200);
                oci_bind_by_name($stmt, ":op_ErrorMsgs", $op_ErrorMsgs, 1000);

                $_resExc = oci_execute($stmt);
                $_err = oci_error($stmt);
                $_resFree = oci_free_statement($stmt);
                if ($_resExc == false) {
                    $op_ErrorCode = "ES";
                    $op_ErrorMsgs = "Error Execute Scan Check Barcode! : " . $_err['message'];
                    wh_log_path(LogFolder, $op_ErrorMsgs);
                    // return
                } else if ($op_ErrorCode == $RetPass) {
                    $DataRows = [
                        'op_Received' => $op_Received,
                        'error_code' => $op_ErrorCode,
                        'error_msge' => $op_ErrorMsgs
                    ];
                    $_log = "pda_qc_output_check_barcode() USER=" . $WK_LOCATER . ",BARCODE=" . $WK_BARCODE . "LOCATOR=" . $WK_LOCATER . "-" . "LOG=" . $op_ErrorCode . "-" . $op_ErrorMsgs . 'Received' . $op_Received;
                    wh_log_path(LogFolder, $op_ErrorMsgs);
                } else {
                    $_log = "Error pda_qc_output_check_barcode() USER=" . $WK_LOCATER . ",BARCODE=" . $WK_BARCODE . "-" . $op_ErrorCode . "-" . $op_ErrorMsgs . 'Received' . $op_Received;
                    wh_log_path(LogFolder, $_log);
                }
                //oci_close($conn);
            } catch (Throwable $e) {
                // Executed only in PHP 7, will not match in PHP 5.x
                $_res = $e;
                $op_ErrorCode = "ES";
                $op_ErrorMsgs = "Error Throwable Scan Check Barcode! " . $_res;
                wh_log_path(LogFolder, $op_ErrorMsgs);
            } catch (Exception $e) {
                // Executed only in PHP 5.x, will not be reached in PHP 7
                $_res = $e;
                $op_ErrorCode = "ES";
                $op_ErrorMsgs = "Error Exception Scan Check Barcode! " . $_res;
                wh_log_path(LogFolder, $op_ErrorMsgs);
            } finally {
                oci_close($conn);
            }
        }
    }

    $DataRows_Count = GET_SUMMARY_PACKING($connectionStr, $WK_PL);
    // มีเจอมากกว่า 1 row  
    if ($DataRows_Count != false) {
        $resultArrayX = array();
        foreach ($DataRows_Count as $rowx) {
            array_push($resultArrayX, $rowx);
        }
        $result_out->cdata = $resultArrayX;
        wh_log_path(LogFolder, 'Dupplicate Job !!');
    }
    // ไม่เจอ  
    else {
        $resultArrayX = array();
        $result_out->cdata = $resultArrayX;
        wh_log_path(LogFolder, 'Job not found !!');
    }


    $DataRows_set = [
        'WK_LOCATER' => $WK_LOCATER,
        'WK_BARCODE' => $WK_BARCODE
    ];

    $DataRows = $DataRows_set;

    $resultArray = array();
    $resultArray2 = array();
    array_push($resultArray, $DataRows);
    array_push($resultArray2, $DataRows_set);
    $result_out->status = (true);
    //$result_out->status = ($op_error_code == $RetPass);
    $result_out->error = $op_ErrorCode;
    $result_out->message = $op_ErrorMsgs;
    $result_out->vdata = $resultArray;
    $result_out->l_data = $resultArray2;
    // return value test
    $result_out->WK_BARCODE = $WK_BARCODE;
    $result_out->input_rec_no = '3324855';
    $result_out->barcode = $WK_BARCODE;
    $result_out->input_barcode_sub = 0;
    $result_out->input_total = 1;
    $result_out->input_roll = 22;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/submit', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $WK_RECEIVED = isset($_REQUEST['locator']) ? $_REQUEST['locator'] : '';
    $connectionStr = isset($_REQUEST['Connection']) ? $_REQUEST['Connection'] : '';
    $WK_USER = isset($_REQUEST['Username']) ? $_REQUEST['Username'] : '';

    $result_out = new stdClass();
    $DataRows = [];
    $op_ErrorCode = '';
    $op_ErrorMsgs = '';

    $conn = GetConnectByLocation($connectionStr);

    $open = false;
    if ($open) {
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
                $command = 'BEGIN PDA_SCAN_RECV_C (
            :WK_RECEIVED,
            :WK_USER, 
            :op_ErrorCode,
            :op_ErrorMsgs
            ); END;';
                $stmt = oci_parse($conn, $command);

                // Bind the input parameter
                oci_bind_by_name($stmt, ":WK_PO", $WK_RECEIVED);
                oci_bind_by_name($stmt, ":WK_LINE", $WK_USER);

                // Bind the output parameter
                oci_bind_by_name($stmt, ":op_Received", $op_Received, 200);
                oci_bind_by_name($stmt, ":op_ErrorCode", $op_ErrorCode, 200);

                $_resExc = oci_execute($stmt);
                $_err = oci_error($stmt);
                $_resFree = oci_free_statement($stmt);
                if ($_resExc == false) {
                    $op_ErrorCode = "ES";
                    $op_ErrorMsgs = "Error Execute Scan Check Barcode! : " . $_err['message'];
                    wh_log_path(LogFolder, $op_ErrorMsgs);
                    // return
                } else if ($op_ErrorCode == $RetPass) {
                    $DataRows = [
                        'error_code' => $op_ErrorCode,
                        'error_msge' => $op_ErrorMsgs
                    ];
                    $_log = "pda_qc_output_check  USER=" . $WK_USER . ",BARCODE=" . $WK_RECEIVED . "connectionStr =" . $connectionStr . "-" . "LOG=" . $op_ErrorCode . "-" . $op_ErrorMsgs;
                    wh_log_path(LogFolder, $op_ErrorMsgs);
                } else {
                    $_log = "Error pda_qc_output_check USER=" . $WK_USER . ",BARCODE=" . $WK_RECEIVED . "-" . $op_ErrorCode . "-" . $op_ErrorMsgs;
                    wh_log_path(LogFolder, $_log);
                }
                //oci_close($conn);
            } catch (Throwable $e) {
                // Executed only in PHP 7, will not match in PHP 5.x
                $_res = $e;
                $op_ErrorCode = "ES";
                $op_ErrorMsgs = "Error Throwable Scan Check Barcode! " . $_res;
                wh_log_path(LogFolder, $op_ErrorMsgs);
            } catch (Exception $e) {
                // Executed only in PHP 5.x, will not be reached in PHP 7
                $_res = $e;
                $op_ErrorCode = "ES";
                $op_ErrorMsgs = "Error Exception Scan Check Barcode! " . $_res;
                wh_log_path(LogFolder, $op_ErrorMsgs);
            } finally {
                oci_close($conn);
            }
        }
    }

    $DataRows_set = [
        'WK_RECEIVED' => $WK_RECEIVED,
        'WK_USER' => $WK_USER
    ];

    $DataRows = $DataRows_set;

    $resultArray = array();
    $resultArray2 = array();
    array_push($resultArray, $DataRows);
    array_push($resultArray2, $DataRows_set);
    $result_out->status = (true);
    //$result_out->status = ($op_error_code == $RetPass);
    $result_out->error = $op_ErrorCode;
    $result_out->message = $op_ErrorMsgs;
    $result_out->vdata = $resultArray;
    $result_out->l_data = $resultArray2;
    // return value test

    $result_out->input_rec_no = '3324855';
    $result_out->input_barcode_sub = 0;
    $result_out->input_total = 1;
    $result_out->input_roll = 22;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/barcode_table', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $RECEIVED_NO = isset($_REQUEST['rec_no']) ? $_REQUEST['rec_no'] : '';

    //$RECEIVED_NO = '381801800039';
    $_sql = "SELECT BARCODE  
    FROM DFWH_WAREHOUSE  
    WHERE RECEIVED_NO = '" . $RECEIVED_NO . "' 
    ORDER BY 1";

    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array();
    $connectionStr = 'HATCH';
    $DataRows = OracleSelectAll($connectionStr, $_sql);
    // มีเจอมากกว่า 1 row  
    if ($DataRows != false) {
        foreach ($DataRows as $row) {
            array_push($resultArray, $row);
        }
        $result_out->vdata = $resultArray;
        $result_out->status = (true);
        $result_out->message = 'Dupplicate Job !!';
        $result_out->error_code = '102';
        $result_out->error = '';
    }
    // ไม่เจอ  
    else {
        $resultArray = array();
        $result_out->vdata = $resultArray;
        $result_out->status = (false);
        $result_out->message = 'Job not found !!';
        $result_out->error_code = '101';
        $result_out->error = 'ES';
    }

    $someArray = [
        [
            "RECEIVED_NO" => $RECEIVED_NO,
        ]
    ];
    // $resultArray = array();
    $resultArray2 = array();
    // array_push($resultArray, 'packing');
    array_push($resultArray2, $someArray);
    $result_out->locator = $RECEIVED_NO;
    $result_out->status = (true);
    // $result_out->vdata = $resultArray;
    $result_out->barcode = $RECEIVED_NO;
    $result_out->input_total = 1;
    $result_out->input_roll = 13;

    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->run(); // สั่งให้ระบบทำงาน
