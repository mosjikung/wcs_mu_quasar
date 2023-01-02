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

define("LogFolder", "LogWSCutProcess");
// $LogFolder = "LogWSLoadingEntry";
// $LogExtensionFile = ".log";
// $LogFileNameFormat = "yyyyMMdd";
$RetPass = "0";

$app = new \Slim\App; // สร้าง object หลักของระบบ

$app->get('/', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $response->getBody()->write("Webservice PDA WCS - CutProcess"); // สร้างคำตอบกลับ
    //wh_log("Webservice PDA WCS - CutProcess");

    return $response; // ส่งคำตอบกลับ
});

//Start Connect Oracle

/**
 * @param  $_query
 * @return array or false
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
        if($_res)
        {
           $nrows = oci_fetch_all($stid, $DataRows, null, null, OCI_FETCHSTATEMENT_BY_ROW); 
        }
        else
        {
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

$app->post('/getmasterjobtype', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';

    $_sql = " select vender_type
              from db_vender
              group by vender_type
              order by vender_type ASC ";

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
        $result_out->error = "";
        $result_out->message = "No data found";
        $result_out->vdata = $resultArray;
    }
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/getmastervender', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';
    $vender_type = isset($_REQUEST['type']) ? $_REQUEST['type'] : '';

    $_sql = " select VENDER_CODE ,  VENDER_NAME
              from db_vender
              WHERE VENDER_TYPE = '$vender_type'
              order by VENDER_NAME ASC ";

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
        $result_out->error = "";
        $result_out->message = "No data found";
        $result_out->vdata = $resultArray;
    }
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/getsono', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';
    $sono = isset($_REQUEST['sono']) ? $_REQUEST['sono'] : 0;

    $_sql = "   SELECT ROWNUM, SO_NO, SO_YEAR, SCHEDULE_ID, MC_LINE, COLOR_FC
                FROM DFIT_SCHEDULE
                WHERE STATUS = 'Y'
                    AND TO_NUMBER('20'||SO_YEAR) > EXTRACT(YEAR FROM SYSDATE) - 3
                    AND SO_NO like '$sono%'
                ORDER BY  SO_NO, SO_YEAR, SCHEDULE_ID ";

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
        $result_out->error = "";
        $result_out->message = "No data found";
        $result_out->vdata = $resultArray;
    }
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

/* Get Barcode Scan */
$app->post('/getbarcodescan', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';
    $jobno = isset($_REQUEST['sono']) ? $_REQUEST['sono'] : 0;

    $_sql = "   select DOC_NO,SO_YEAR,SO_NO,BARCODE,SO_COLOR,SO_SIZE,BOX_NO as QTY 
                from dfit_ep_detail  
                where bu_code ='NYP'
                and update_date >= to_date('24/04/2020','DD/MM/RRRR')  
                and  doc_no  ='$jobno'
                order by update_date DESC  ";

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
        $result_out->error = "";
        $result_out->message = "No data found";
        $result_out->vdata = $resultArray;
    }
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/genjob', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';

    $sono = isset($_REQUEST['sono']) ? $_REQUEST['sono'] : 0;
    $soyear = isset($_REQUEST['soyear']) ? $_REQUEST['soyear'] : 0;
    $jobtype = isset($_REQUEST['jobtype']) ? $_REQUEST['jobtype'] : 0;
    $venderid = isset($_REQUEST['venderid']) ? $_REQUEST['venderid'] : 0;
    $mcline = isset($_REQUEST['mcline']) ? $_REQUEST['mcline'] : 0;
    $scheduleid = isset($_REQUEST['scheduleid']) ? $_REQUEST['scheduleid'] : 0;

    $result_out = new stdClass();
    $DataRows = [];
    $op_error_code = '';
    $op_error_msge = '';

    if ($soyear == '' || $sono == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "ES";
        $result_out->message = "Error Entry SO No. or  SO Year ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY SO ! USER=" . $uid;
        wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($jobtype == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "ES";
        $result_out->message = "Error Entry Jobtype ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY Jobtype ! USER=" . $uid;
        wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($venderid == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "Error Entry Vender ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY Vender ! USER=" . $uid;
        wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($mcline == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "Error Entry Line ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY Line ! USER=" . $uid;
        wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($scheduleid == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "Error Entry Schedule ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY Schedule ! USER=" . $uid;
        wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    /*
    CREATE PROCEDURE pda_cutprocess_genjob (
    ipso_no         IN       NUMBER,
    ipso_year       IN       NUMBER,
    ipstep_name     IN       VARCHAR2,
    ipvender_id     IN       VARCHAR2,
    ipupdate_by     IN       VARCHAR2,
    ipmc_line       IN       VARCHAR2,
    ipschedule_id   IN       VARCHAR2,
    opdoc_no        OUT      VARCHAR2,
    operror_code    OUT      NUMBER,
    operror_msge    OUT      VARCHAR2
    )
     */

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
            $command = 'BEGIN pda_cutprocess_genjob (
                                      :ipso_no, :ipso_year, :ipstep_name, :ipvender_id, :ipupdate_by,
                                      :ipmc_line, :ipschedule_id,
                                      :opdoc_no,
                                      :op_error_code, 
                                      :op_error_msge
                        ); END;';
            $stmt = oci_parse($conn, $command);

            // Bind the input parameter
            oci_bind_by_name($stmt, ":ipso_no", $sono);
            oci_bind_by_name($stmt, ":ipso_year", $soyear);
            oci_bind_by_name($stmt, ":ipstep_name", $jobtype);
            oci_bind_by_name($stmt, ":ipvender_id", $venderid);
            oci_bind_by_name($stmt, ":ipupdate_by", $uid);
            oci_bind_by_name($stmt, ":ipmc_line", $mcline);
            oci_bind_by_name($stmt, ":ipschedule_id", $scheduleid);

            // Bind the output parameter
            oci_bind_by_name($stmt, ":op_opdoc_no", $op_opdoc_no, 100);
            oci_bind_by_name($stmt, ":op_error_code", $op_error_code, 100);
            oci_bind_by_name($stmt, ":op_error_msge", $op_error_msge, 200);

            $_resExc = oci_execute($stmt);
            $_err = oci_error($stmt);
            $_resFree = oci_free_statement($stmt);
            if($_resExc == false)
            {   
                $op_error_code = "ES";
                $op_error_msge = "Error Execute Gen Job ! : ". $_err['message'];       
                wh_log_path(LogFolder, $op_error_msge);        
                // return 
            }else if ($op_error_code == $RetPass)
            {
                $_log = "USER=".$uid." - GEN_JOB " . $op_opdoc_no;
                wh_log_path(LogFolder, $_log);
                 $DataRows = [
                    'job_no' => $op_opdoc_no,
                    'error_code' => $op_error_code,
                    'error_msge' => $op_error_msge,
                ];
            }else{
                $_log = "Error pda_cutprocess_genjob() USER=".$uid." - ".$op_error_code."-".$op_error_msge;
                wh_log_path(LogFolder, $_log); 
            }

            $_resExc = oci_execute($stmt);
            //print "$p2\n";   // prints 16
            $_resFree = oci_free_statement($stmt);
            if ($_resExc == false) {
                $_err = oci_error($command);

                $resultArray = array();
                $result_out->status = false;
                $result_out->error = "Error Gen Job ! " . $_err;
                $result_out->code = "ES";
                $result_out->message = "Error Gen Job ! ";
                $result_out->vdata = $resultArray;
                $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

                wh_log_path(LogFolder, $result_out->error);

                return $response; // ส่งคำตอบกลับ
            } else {
                $log_barcode_confirm = "GEN_JOB " . $op_opdoc_no;
                wh_log_path(LogFolder, $log_barcode_confirm);
            }

            //oci_close($conn);
        } catch (Throwable $e) {
            // Executed only in PHP 7, will not match in PHP 5.x
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Exception Gen Job ! ". $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } catch (Exception $e) {
            // Executed only in PHP 5.x, will not be reached in PHP 7
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Exception Gen Job ! ". $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } finally {
            oci_close($conn);
        }

    }

    // //Test
    // $op_error_code = "0";
    // $DataRows = [
    //     'job_no' => "DOC999999",
    //     'error_code' => "",
    //     'error_msge' => "",
    // ];

    $resultArray = array();
    array_push($resultArray, $DataRows);
    $result_out->status = ($op_error_code == $RetPass);
    $result_out->error = $op_error_code;
    $result_out->message = $op_error_msge;    
    $result_out->vdata = $resultArray;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/scanbarcode', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';

    $jobno = isset($_REQUEST['jobno']) ? $_REQUEST['jobno'] : 0;
    $sono = isset($_REQUEST['sono']) ? $_REQUEST['sono'] : 0;
    $soyear = isset($_REQUEST['soyear']) ? $_REQUEST['soyear'] : 0;
    $jobtype = isset($_REQUEST['jobtype']) ? $_REQUEST['jobtype'] : '';
    $venderid = isset($_REQUEST['venderid']) ? $_REQUEST['venderid'] : '';
    $mcline = isset($_REQUEST['mcline']) ? $_REQUEST['mcline'] : '';
    $scheduleid = isset($_REQUEST['scheduleid']) ? $_REQUEST['scheduleid'] : '';
    $barcode = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';

    $result_out = new stdClass();
    $DataRows = [];
    $op_error_code ='';
    $op_error_msge ='';

    if ($barcode == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "PLEASE INSERT BARCODE ! ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY BARCODE ! USER=" . $uid . ",BARCODE=" . $barcode;
        wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($soyear == '' || $sono == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "Error Entry SO No. or  SO Year ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        //$strlog = "ENTRY SO ! USER=" . $uid . ",BARCODE=" . $barcode;
        //wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($jobtype == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "Error Entry Jobtype ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        //$strlog = "ENTRY Jobtype ! USER=" . $uid;
        //wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($venderid == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "Error Entry Vender ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        //$strlog = "ENTRY Vender ! USER=" . $uid;
        //wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($mcline == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "Error Entry Line ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        //$strlog = "ENTRY Line ! USER=" . $uid;
        //wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($scheduleid == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "Error Entry Schedule ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        //$strlog = "ENTRY Schedule ! USER=" . $uid;
        //wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    /*
    pda_cutprocess_cfmbarcode (
    ipso_no         IN       VARCHAR2,
    ipso_year       IN       VARCHAR2,
    ipstep_name     IN       VARCHAR2,
    ipvender_id     IN       VARCHAR2,
    ipupdate_by     IN       VARCHAR2,
    ipmc_line       IN       VARCHAR2,
    ipschedule_id   IN       VARCHAR2,
    ipdoc_no        IN       VARCHAR2,
    ipbarcode       IN       VARCHAR2,
    opcolor         OUT      VARCHAR2,
    opsize          OUT      VARCHAR2,
    opqty           OUT      NUMBER,
    operror_code    OUT      NUMBER,
    operror_msge    OUT      VARCHAR2
     */

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
            $command = 'BEGIN pda_cutprocess_cfmbarcode(
                                      :ipso_no, :ipso_year, :ipstep_name, :ipvender_id, :ipupdate_by,
                                       :ipmc_line, :ipschedule_id, :ipdoc_no, :ipbarcode,
                                        :opcolor, :opsize, :opqty,
                                        :op_error_code, :op_error_msge
                        ); END;';
            $stmt = oci_parse($conn, $command);

            // Bind the input parameter
            oci_bind_by_name($stmt, ":ip_empuser", $uid);
            oci_bind_by_name($stmt, ":ipso_no", $sono);
            oci_bind_by_name($stmt, ":ipso_year", $soyear);
            oci_bind_by_name($stmt, ":ipstep_name", $jobtype);
            oci_bind_by_name($stmt, ":ipvender_id", $venderid);
            oci_bind_by_name($stmt, ":ipupdate_by", $uid);
            oci_bind_by_name($stmt, ":ipmc_line", $mcline);
            oci_bind_by_name($stmt, ":ipschedule_id", $scheduleid);
            oci_bind_by_name($stmt, ":ipdoc_no", $jobno);
            oci_bind_by_name($stmt, ":ipbarcode", $barcode);

            // Bind the output parameter
            oci_bind_by_name($stmt, ":op_error_code", $op_error_code, 100);
            oci_bind_by_name($stmt, ":op_error_msge", $op_error_msge, 200);

            $_resExc = oci_execute($stmt);
            $_err = oci_error($stmt);
            $_resFree = oci_free_statement($stmt);
            if($_resExc == false)
            {   
                $op_error_code = "ES";
                $op_error_msge = "Error Execute Scan Barcode! : ". $_err['message'];       
                wh_log_path(LogFolder, $op_error_msge);        
                // return 
            }else if ($op_error_code == $RetPass)
            {
                $DataRows = [
                    'error_code' => $op_error_code,
                    'error_msge' => $op_error_msge,
                ];
                $_log = " Scan Barcode : userid=".$uid.", barcode=".$barcode;
                wh_log_path(LogFolder, $_log);
            }else{
                $_log = "Error pda_cutprocess_cfmbarcode() USER=".$uid.",BARCODE=".$barcode." - ".$op_error_code."-".$op_error_msge;
                wh_log_path(LogFolder, $_log); 
            }
            //oci_close($conn);
        } catch (Throwable $e) {
            // Executed only in PHP 7, will not match in PHP 5.x
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Throwable Scan Barcode! ". $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } catch (Exception $e) {
            // Executed only in PHP 5.x, will not be reached in PHP 7
            $_res = $e;
            $op_error_code = "ES";
            $op_error_msge = "Error Exception Scan Barcode! ". $_res;
            wh_log_path(LogFolder, $op_error_msge);
        } finally {
            oci_close($conn);
        }
    }

    //Test
    // $op_error_code = "0";
    // $DataRows = [
    //     'error_code' => $op_error_code,
    //     'error_msge' => $op_error_msge,
    //     'barcode' => $barcode,
    //     'color' => "red",
    //     'size' => "XXX",
    //     'qty' => 10,
    // ];

    $resultArray = array();
    array_push($resultArray, $DataRows);
    $result_out->status = ($op_error_code == $RetPass);
    $result_out->error = $op_error_code;
    $result_out->message = $op_error_msge;    
    $result_out->vdata = $resultArray;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/deletebarcode', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';

    $jobno = isset($_REQUEST['jobno']) ? $_REQUEST['jobno'] : 0;
    $sono = isset($_REQUEST['sono']) ? $_REQUEST['sono'] : 0;
    $soyear = isset($_REQUEST['soyear']) ? $_REQUEST['soyear'] : 0;
    $barcode = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';

    $result_out = new stdClass();
    $DataRows = [];
    $op_error_code ='';
    $op_error_msge ='';

    if ($barcode == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "PLEASE INSERT BARCODE ! ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY BARCODE ! USER=" . $uid . ",BARCODE=" . $barcode;
        //wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }

    if ($soyear == '' || $sono == '') {
        $resultArray = array();
        $result_out->status = false;
        $result_out->error = "";
        $result_out->code = "ES";
        $result_out->message = "Error Entry SO No. or  SO Year ! ";
        $result_out->vdata = $resultArray;
        $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

        $strlog = "ENTRY SO ! USER=" . $uid . ",BARCODE=" . $barcode;
        //wh_log_path(LogFolder, $strlog);

        return $response; // ส่งคำตอบกลับ
    }


    //Test
    $op_error_code = "1";
    $op_error_msge = "Error Delete !";
    $DataRows = [
        'error_code' => $op_error_code,
        'error_msge' => $op_error_msge,
        'barcode' => $barcode
    ];

    $resultArray = array();
    array_push($resultArray, $DataRows);
    $result_out->status = ($op_error_code == $RetPass);
    $result_out->error = $op_error_code;
    $result_out->message = $op_error_msge;    
    $result_out->vdata = $resultArray;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/confirmjob', function (Request $request, Response $response)  use($RetPass)  { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';

    $jobno = isset($_REQUEST['jobno']) ? $_REQUEST['jobno'] : 0;
    $sono = isset($_REQUEST['sono']) ? $_REQUEST['sono'] : 0;
    $soyear = isset($_REQUEST['soyear']) ? $_REQUEST['soyear'] : 0;
    $jobtype = isset($_REQUEST['jobtype']) ? $_REQUEST['jobtype'] : '';
    $venderid = isset($_REQUEST['venderid']) ? $_REQUEST['venderid'] : '';
    $mcline = isset($_REQUEST['mcline']) ? $_REQUEST['mcline'] : '';
    $scheduleid = isset($_REQUEST['scheduleid']) ? $_REQUEST['scheduleid'] : '';
    
    $result_out = new stdClass();
    $DataRows = [];
    $op_error_code ='';
    $op_error_msge ='';

    //Pass
    $op_error_code ='0';
    $DataRows = [
        'jobno' => $jobno,
        'error_code' => "",
        'error_msge' => "",
    ];
    wh_log_path(LogFolder, "Confirm Job : userid=$uid , jobno=$jobno");

    $resultArray = array();
    array_push($resultArray, $DataRows);
    $result_out->status = ($op_error_code == $RetPass);
    $result_out->error = $op_error_code;
    $result_out->message = $op_error_msge;    
    $result_out->vdata = $resultArray;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

function getRunning()
{
    /*
begin
select max(nvl(job_no,0))+100001
into get_running
from AD_TRANSACTION_HEAD
where job_year = to_char(sysdate,'YY');
exception when no_data_found then
get_running := '100001';
end;
if (get_running is null) then
get_running := '100001';
end if;

 */
}

$app->run(); // สั่งให้ระบบทำงาน

/*
// Define MYDB connection string as described in tnsnames.ora
define("MYDB","(DESCRIPTION=(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = JXYX.com)(PORT = 1521)))(CONNECT_DATA=(SID=DHSJKS)))");
// Connect to database
$conn = oci_connect("XXXXXX","XXXXXXXX",MYDB);
// Through error if not connected
if (!$conn) {
$e = oci_error();
trigger_error(htmlentities($e['message'], ENT_QUOTES), E_USER_ERROR);
}
// Bind the your input and output parameters to PHP variables
$stmt = oci_parse($conn,'BEGIN SP_GET_MY_DATA(:POP, :SEG, :DUR, :VIEW, :PAGE, :OUTPUT_CUR); END;');
oci_bind_by_name($stmt,':POP',$pop);
oci_bind_by_name($stmt,':SEG',$seg);
oci_bind_by_name($stmt,':DUR',$dur);
oci_bind_by_name($stmt,':VIEW',$view);
oci_bind_by_name($stmt,':PAGE',$page);
// Declare your cursor
$OUTPUT_CUR = oci_new_cursor($conn);
oci_bind_by_name($stmt,":OUTPUT_CUR", $OUTPUT_CUR, -1, OCI_B_CURSOR);
// Execute statement
oci_execute($stmt);
// Execute the cursor
oci_execute($OUTPUT_CUR);
// Fetch results
while ($data = oci_fetch_assoc($OUTPUT_CUR)) {
print_r($data);
}
 */
