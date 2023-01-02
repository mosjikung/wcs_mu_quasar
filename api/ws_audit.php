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


// Check Role
$app->post('/ADUIT_MAIN', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $locator = isset($_REQUEST['locator']) ? $_REQUEST['locator'] : '';
    $barcode = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';

    /* for ($x = 0; $x <= 10; $x++) {
        echo "The number is: $x <br>";
    } */
    $someArray = [
        [
            "id" => 1,
            "name" =>'name1',
            "uid" => $uid,
            "locator" => $locator,
            "barcode" => $barcode,
        ]
    ];

    $result_out = new stdClass();
    $resultArray = array();
    $resultArray2 = array();
    array_push($resultArray, 'xxx');
    array_push($resultArray2, $someArray);
    $result_out->status = (true);
    $result_out->error = 'xxzzzz';
    $result_out->message = '';
    $result_out->barcode = $barcode;
    $result_out->locator = $locator;
    $result_out->vdata = $resultArray;
    $result_out->l_data = $resultArray2;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->post('/ADUIT_COUNT', function (Request $request, Response $response) use ($RetPass) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $uid = isset($_REQUEST['uid']) ? $_REQUEST['uid'] : '';
    $locator = isset($_REQUEST['locator']) ? $_REQUEST['locator'] : '';
    $barcode = isset($_REQUEST['barcode']) ? $_REQUEST['barcode'] : '';

    /* for ($x = 0; $x <= 10; $x++) {
        echo "The number is: $x <br>";
    } */
    $someArray = [
        [
            "id" => 1,
            "name" =>'count1',
            "uid" => $uid,
            "locator" => $locator,
            "barcode" => $barcode,
        ]
    ];
    $result_out = new stdClass();
    $resultArray = array();
    $resultArray2 = array();
    array_push($resultArray, 'xxx');
    array_push($resultArray2, $someArray);
    $result_out->status = (true);
    $result_out->error = 'xxzzzz';
    $result_out->message = '';
    $result_out->barcode = $barcode;
    $result_out->locator = $locator;
    $result_out->vdata = $resultArray;
    $result_out->l_data = $resultArray2;
    $response->getBody()->write(json_encode($result_out)); // สร้างคำตอบกลับ

    return $response; // ส่งคำตอบกลับ
});

$app->run();
