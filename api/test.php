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

// Connect Database
// define("BASE_NYG1HO_HOST", "(DESCRIPTION=(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.82)(PORT = 1521)))(CONNECT_DATA=(SID=NYTG)))");
// define("BASE_NYG1HO_USER", "NYG1");
// define("BASE_NYG1HO_PASSWORD", "NYG1");

date_default_timezone_set("Asia/Bangkok");

$app = new \Slim\App; // สร้าง object หลักของระบบ

$app->get('/fileperms', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    // $response->getBody()->write("Webservice PDA for NYG") . " root "; // สร้างคำตอบกลับ

    $fn = './log';
    $fp = fileperms($fn);

    echo substr(sprintf('%o', $fp), -4) . '<br/>'; //0666

    chmod($fn, 0755);

    $msg = is_readable($fn) ? $msg = 'File is readable'
    : $msg = 'File is not readable';
    echo $msg . '<br/>';

    $msg = is_writable($fn) ? $msg = 'File is writable'
    : $msg = 'File is not writable';

    echo $msg . '<br/>';

    $msg = is_executable($fn) ? $msg = 'File is executable'
    : $msg = 'File is not executable';

    echo $msg . '<br/>';

    // $log_filename = getcwd() . "/log_test";
    // try {
    //     mkdir($log_filename, 0777, true);
    // } catch (\Throwable $th) {
    //     $response->getBody()->write("Create folder Throwable".$log_filename." ".$th); // สร้างคำตอบกลับ
    // } catch (Exception $th) {
    //     $response->getBody()->write("Create folder Exception".$log_filename." ".$th); // สร้างคำตอบกลับ
    // }

    // $response->getBody()->write($log_filename); // สร้างคำตอบกลับ
    //wh_log("Webservice PDA NYG System.");

    return $response; // ส่งคำตอบกลับ
});

$app->get('/get_client_ip', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
    // $response->getBody()->write("Webservice PDA for NYG") . " root "; // สร้างคำตอบกลับ

    $msg = get_client_ip();
    echo $msg . '<br/>';

    // $log_filename = getcwd() . "/log_test";
    // try {
    //     mkdir($log_filename, 0777, true);
    // } catch (\Throwable $th) {
    //     $response->getBody()->write("Create folder Throwable".$log_filename." ".$th); // สร้างคำตอบกลับ
    // } catch (Exception $th) {
    //     $response->getBody()->write("Create folder Exception".$log_filename." ".$th); // สร้างคำตอบกลับ
    // }

    // $response->getBody()->write($log_filename); // สร้างคำตอบกลับ
    //wh_log("Webservice PDA NYG System.");

    return $response; // ส่งคำตอบกลับ
});

// function getPHPBitVersion()
// {
//     switch (PHP_INT_SIZE) {
//         case 4:
//             return '32-bit version of PHP';
//             break;
//         case 8:
//             return '64-bit version of PHP';
//             break;
//         default:
//             return 'PHP_INT_SIZE is ' . PHP_INT_SIZE;
//     }
// }

// /**
//  * @param $log_msg
//  */
// function wh_log($log_msg)
// {
// // To get your current working directory: getcwd() (documentation)
//     // To get the document root directory: $_SERVER['DOCUMENT_ROOT']
//     // To get the filename of the current script: $_SERVER['SCRIPT_FILENAME']
//     $log_filename = getcwd() . "/log";

//     $log_time = date('Y-m-d h:i:sa');
//     $log_head = "***** Start Log For Day : '" . $log_time . "'****** ";
//     $log_foot = "***** END Log For Day : '" . $log_time . "'***** \r\n";
//     $log_msg = $log_head . "\r\n" . $log_msg . "\r\n" . $log_foot;

//     if (!file_exists($log_filename)) {
//         // create directory/folder uploads.
//         mkdir($log_filename, 0777, true);
//     }
//     $log_file_data = $log_filename . '/log_' . date('Y-M-d') . '.log';
//     file_put_contents($log_file_data, $log_msg . "\r\n", FILE_APPEND);
// }

// //Start Connect Oracle

// /**
//  * @param  $_query
//  * @return array
//  */
// function OracleSelectAll($_query)
// {
//     $conn = oci_connect(BASE_NYG1HO_USER, BASE_NYG1HO_PASSWORD, BASE_NYG1HO_HOST, 'AL32UTF8');
//     $DataRows = array();

//     if (!$conn) {
//         $m = oci_error();
//         wh_log($m['message']);

//         return $m['message'];
//         //echo $m['message'], "\n";
//         //exit;
//     }

//     try {

//         $stid = oci_parse($conn, $_query);
//         oci_execute($stid);
//         $nrows = oci_fetch_all($stid, $DataRows, null, null, OCI_FETCHSTATEMENT_BY_ROW);
//     } catch (Throwable $e) {
//         // Executed only in PHP 7, will not match in PHP 5.x
//         $_res = $e;
//         wh_log($_res);
//     } catch (Exception $e) {
//         // Executed only in PHP 5.x, will not be reached in PHP 7
//         $_res = $e;
//         wh_log($_res);

//     } finally {
//         oci_close($conn);
//     }

//     return $DataRows;
// }

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

// //End Connect Oracle

// /// Test
// $app->post('/test', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
//     $lang = isset($_REQUEST['lang']) ? $_REQUEST['lang'] : '';
//     $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
//     $so_no = isset($_REQUEST['so_no']) ? $_REQUEST['so_no'] : '';
//     $so_year = isset($_REQUEST['so_year']) ? $_REQUEST['so_year'] : '';
//     $barcode = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';

//     $result_out = new stdClass();

//     // Connect to database
//     $conn = oci_connect(BASE_NYG1HO_USER, BASE_NYG1HO_PASSWORD, BASE_NYG1HO_HOST);
//     // Through error if not connected
//     if (!$conn) {
//         // $e = oci_error();
//         //trigger_error(htmlentities($e['message'], ENT_QUOTES), E_USER_ERROR);
//         $resultArray = array();
//         $result_out->status = false;
//         $result_out->error = "Error Connect Database! ";
//         $result_out->message = "";
//         $result_out->vdata = $resultArray;
//         $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

//         return $response; // ส่งคำตอบกลับ

//     } else {
//         // Bind the your input and output parameters to PHP variables
//         $sql = 'BEGIN pda_loaing_entry_cfm_barcode(
//                             :ip_empuser, :ip_so_no, :ip_so_year, :ip_barcode,
//                             :op_error_code, :op_error_msge,
//                             :op_barcode, :op_so_no, :op_so_year, :op_customer_id, :op_customer_name,
//                             :op_plan_id, :op_style, :op_line, :op_color, :op_size, :op_qty
//                 ); END;';
//         $stmt = oci_parse($conn, $sql);

//         // Bind the output parameter

//         oci_bind_by_name($stmt, ":ip_empuser", $uid);
//         oci_bind_by_name($stmt, ":ip_so_no", $so_no);
//         oci_bind_by_name($stmt, ":ip_so_year", $so_year);
//         oci_bind_by_name($stmt, ":ip_barcode", $barcode);

//         oci_bind_by_name($stmt, ":op_error_code", $op_error_code, 100);
//         oci_bind_by_name($stmt, ":op_error_msge", $op_error_msge, 200);

//         oci_bind_by_name($stmt, ":op_barcode", $op_barcode, 200);
//         oci_bind_by_name($stmt, ":op_so_no", $op_so_no, 40);
//         oci_bind_by_name($stmt, ":op_so_year", $op_so_year, 40);
//         oci_bind_by_name($stmt, ":op_customer_id", $op_customer_id, 200);
//         oci_bind_by_name($stmt, ":op_customer_name", $op_customer_name, 200);
//         oci_bind_by_name($stmt, ":op_plan_id", $op_plan_id, 50);
//         oci_bind_by_name($stmt, ":op_style", $op_style, 50);
//         oci_bind_by_name($stmt, ":op_line", $op_line, 50);
//         oci_bind_by_name($stmt, ":op_color", $op_color, 50);
//         oci_bind_by_name($stmt, ":op_size", $op_size, 40);
//         oci_bind_by_name($stmt, ":op_qty", $op_qty, 40);

//         oci_execute($stmt);
//         //print "$p2\n";   // prints 16
//         oci_free_statement($stmt);
//         oci_close($conn);

//     }

//     $DataRows = [
//         'error_code' => $op_error_code,
//         'error_msge' => $op_error_msge,
//     ];
//     $resultArray = array();
//     //$DataRows["op_error_code"]->$op_error_code = $op_error_code;
//     //$DataRows["op_error_msge"]->$op_error_msge = $op_error_msge;
//     array_push($resultArray, $DataRows);
//     $result_out->status = true;
//     $result_out->error = "";
//     $result_out->message = "";
//     $result_out->vdata = $resultArray;

//     // Call Prc

//     // if ($DataRows != false) {
//     //     foreach ($DataRows as $row) {
//     //         //echo $row['VERSION'] . '<br>';
//     //         array_push($resultArray, $row);
//     //     }
//     //     $result_out->status = true;
//     //     $result_out->error = "";
//     //     $result_out->message = "";
//     //     $result_out->vdata = $resultArray;
//     // } else {
//     //     $resultArray = array();
//     //     $result_out->status = true;
//     //     $result_out->error = "";
//     //     $result_out->message = "No data found";
//     //     $result_out->vdata = $resultArray;
//     // }
//     $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

//     return $response; // ส่งคำตอบกลับ
// });

// function oci_procedure_call($package_name, $proc_name, $param)
// {
//     $sql = "begin $package_name.$proc_name(";

//     foreach ($param as $p) {
//         $sql .= $p['name'] . ',';
//     }

//     $sql = trim($sql, ',') . '); end;';

//     $stmt = oci_parse($this->db->conn_id, $sql);

//     foreach ($param as $p) {
//         oci_bind_by_name($stmt, $p['name'], $p['value'], $p['length'], $p['type']);
//     }

//     try {
//         $r = oci_execute($stmt);
//     } catch (Exception $e) {
//         //log_message('info', print_r($e, true));
//     }

//     if (!$r) {
//         $e = oci_error($this->db->conn_id);
//         return $e['message'];
//     } else {
//         return true;
//     }
// }

$app->run(); // สั่งให้ระบบทำงาน
