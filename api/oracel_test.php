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
?>
<html>
<head>
<title>ThaiCreate.Com PHP & Oracle Tutorial</title>
</head>
<body>
<?php
$conn = GetConnectByLocation($_baselocation);
$strSQL = "SELECT * FROM DB_USER";
$objParse = oci_parse($conn, $strSQL);
oci_execute ($objParse,OCI_DEFAULT);
?>
<table width="600" border="1">
  <tr>
    <th width="91"> <div align="center">CustomerID </div></th>
    <th width="98"> <div align="center">Name </div></th>
    <th width="198"> <div align="center">Email </div></th>
    <th width="97"> <div align="center">CountryCode </div></th>
    <th width="59"> <div align="center">Budget </div></th>
    <th width="71"> <div align="center">Used </div></th>
  </tr>
<?php
while($objResult = oci_fetch_array($objParse,OCI_BOTH))
{
?>
  <tr>
    <td><div align="center"><?php echo $objResult["EMP_NAME"];?></div></td>
    <td><?php echo $objResult["EMP_NAME"];?></td>
    <td><?php echo $objResult["EMP_NAME"];?></td>
    <td><div align="center"><?php echo $objResult["EMP_CODE"];?></div></td>
    <td align="right"><?php echo $objResult["EMP_CODE"];?></td>
    <td align="right"><?php echo $objResult["EMP_CODE"];?></td>
  </tr>
<?php
}
?>
</table>
<?php
oci_close($objConnect);
?>
</body>
</html>