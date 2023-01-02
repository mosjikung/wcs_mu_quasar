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

define("LogFolder", "LogWSMain");

$app = new \Slim\App; // สร้าง object หลักของระบบ

$app->post('/', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $response->getBody()->write("Webservice PDA WCS System"); // สร้างคำตอบกลับ
    //wh_log("Webservice PDA WCS System.");
    wh_log_path(LogFolder, "Webservice PDA WCS System.");

    return $response; // ส่งคำตอบกลับ
});

$app->post('/getbaselocation', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $DataRows = BASE_LOCATION;
    $resultArray = array();
    foreach ($DataRows as $row) {
        $_location = new stdClass();
        $_location->LOCATION = $row;
        array_push($resultArray, $_location);
    }
    $result_out = new stdClass();
    $result_out->status = true;
    $result_out->error = "";
    $result_out->message = "";
    $result_out->vdata = $resultArray;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

//Start Connect MySQL

/**
 * @param  $array_query
 * @return mixed
 */
function MySqliMoreQuery($array_query)
{
    $_res = "";

    // $conn = MySqliConnect_NYV();
    // $conn->autocommit(false);

    // try {

    //     // Check connection
    //     if ($conn->connect_errno) {
    //         //error_log(sprintf("Connect failed: %s\n", mysqli_connect_error()), 0);
    //         $_res = sprintf("Connect failed: %s\n", mysqli_connect_error());
    //         wh_log($_res);

    //         return $_res;
    //     }

    //     if (count($array_query) == 0 || $array_query == null) {
    //         $_res = "Query is null";

    //         return $_res;
    //     }

    //     foreach ($array_query as $_query) {
    //         $conn->query($_query);
    //         ///if( $conn->affected_rows < 0 && $conn->insert_id == 0)
    //         if ($conn->error != "") {
    //             $_res = $conn->error;
    //             wh_log($_res);
    //             break;
    //         }
    //     }
    //     if ($_res == "") {
    //         $conn->commit();
    //     } else {
    //         $conn->rollback();
    //     }
    // } catch (Throwable $e) {
    //     // Executed only in PHP 7, will not match in PHP 5.x
    //     $conn->rollback();
    //     $_res = $e;
    //     wh_log($_res);
    // } catch (Exception $e) {
    //     // Executed only in PHP 5.x, will not be reached in PHP 7
    //     $conn->rollback();
    //     $_res = $e;
    //     wh_log($_res);

    // } finally {
    //     $conn->close();

    //     return $_res;
    // }
}

//End Connect MySQL

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

// // /**
// //  * @param  $_query
// //  * @return mixed
// //  */
// /**
//  * @param  $_query
//  * @return mixed
//  */
// function OracleSelectAll_Class($_query)
// {
//     $ora = new ORACLE(); //::CreateInstance("XE", "olton", "yfnfkmz", "RUSSIAN_CIS.AL32UTF8");
//     $ora->Connect(BASE_NYG1HO_HOST, "NYG1", "NYG1");

//     $ora->SetFetchMode(OCI_ASSOC);
//     $ora->SetAutoCommit(true);

//     $DataRows = array();

//     try {

//         // Check connection
//         // if ($conn->connect_errno) {
//         //     //error_log(sprintf("Connect failed: %s\n", mysqli_connect_error()), 0);
//         //     $_res = sprintf("Connect failed: %s\n", mysqli_connect_error());
//         //     wh_log($_res);

//         //     return $_res;
//         // }

//         // if ($_query == "" || $_query == null) {
//         //     $_res = "Query is null";

//         //     return $_res;
//         // }

//         // $DataRows = $conn->query($_query)->fetch_all(MYSQLI_ASSOC);

//         $RunSelect = $ora->Select($_query);
//         $DataRows = $ora->FetchAll($RunSelect);

//     } catch (Throwable $e) {
//         // Executed only in PHP 7, will not match in PHP 5.x
//         $_res = $e;
//         wh_log($_res);
//     } catch (Exception $e) {
//         // Executed only in PHP 5.x, will not be reached in PHP 7
//         $_res = $e;
//         wh_log($_res);

//     } finally {
//         $ora->Bye();
//     }

//     return $DataRows;
// }

//End Connect Oracle

//Start Main App

/**
 * @return mixed
 */
$app->post('/signin', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
    $ulocation = isset($_REQUEST['ulocation']) ? $_REQUEST['ulocation'] : '';
    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $pwd = isset($_REQUEST['pwd']) ? $_REQUEST['pwd'] : '';

    $_sql = " SELECT EMP_NAME USER_NAME,EMP_CODE USER_ID
              FROM DB_USER
              WHERE UPPER(ACTIVE)='Y'
                AND EMP_CODE ='$uid'
                AND PASSWORD = '$pwd'
              ORDER BY  UPD_DATE DESC ";

    $result_out = new stdClass();
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

        $_sql = " SELECT USER_ID, USER_NAME
              FROM SU_USER
              WHERE UPPER(ACTIVE)='Y'
                AND USER_ID ='$uid'
                AND USER_PASSWORD = '$pwd'
              ORDER BY  UPD_DATE DESC ";
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
            $result_out->message = "Invalid user/password !";
            $result_out->vdata = $resultArray;
        }
    }

    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ
    wh_log_path(LogFolder, $uid . " message :" . $result_out->message);

    return $response; // ส่งคำตอบกลับ
});

$app->run(); // สั่งให้ระบบทำงาน
