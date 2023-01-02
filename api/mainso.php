<?php
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: access");
header("Access-Control-Allow-Methods: GET,POST");
header("Access-Control-Allow-Credentials: true");
header('Content-Type: application/json;charset=utf-8');
require_once __DIR__ . '/../api/web_config.php';

set_time_limit ( 60000 );

use \Psr\Http\Message\ResponseInterface as Response; // ไลบราลี้สำหรับจัดการคำร้องขอ
use \Psr\Http\Message\ServerRequestInterface as Request; // ไลบราลี้สำหรับจัดการคำตอบกลับ

require './vendor/autoload.php'; // ดึงไฟ์ autoload.php เข้ามา
//include_once './class.oracle.php'; // Class Connect Oracle
include_once './util.php'; // ดึงไฟ์ util.php เข้ามา
include_once './web_config.php';
$app = new \Slim\App; // สร้าง object หลักของระบบ

date_default_timezone_set("Asia/Bangkok");

function ConnectDbAll($_sql){
  $DataRows = array();
  $conn = ConnectedDBSO();
  if(!$conn){
    $_err = oci_error();
    echo $_err;
  }else{
    $objParse = oci_parse($conn,$_sql);
    $objEx = oci_execute($objParse);
    if($objEx){
      $objResult = oci_fetch_all($objParse,$DataRows,null,null, OCI_FETCHSTATEMENT_BY_ROW);
    }else{
      echo "Connect Data Base error";
    }
  }
  oci_close($conn);
  return $DataRows;
 
}
function ConnectDbnoAll($_sql){
  $DataRows = array();
  $conn = ConnectedDBSO();
  if(!$conn){
    $_err = oci_error();
    echo $_err;
  }else{
    $objParse = oci_parse($conn,$_sql);
    $objEx = oci_execute($objParse);
    if($objEx){
      oci_commit($conn); 
    }else{
      echo "Connect Data Base error";
    }
  }
  oci_close($conn);
  return $DataRows;
  
}

 $app->post('/checkso', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  
 
   $_sql = "SELECT SO_NO_DOC, SO_NO, SO_YEAR, QTY, CUST_NAME, STYLE_REF, 
  SHIPMENT_DATE, GMT_TYPE,FABRIC_DESCRIPTION FABRIC_TYPE
   FROM oe_ME_ORDER WHERE SO_YEAR>='".$so_year."' AND SO_NO='".$so_no."'";
  $result_out = new stdClass();
  $DataRows = [];
  $resultArray = array(); //data
  $DataRows = ConnectDbAll($_sql);

  if($DataRows != false){
  foreach ($DataRows as $row){
    array_push($resultArray,$row);
  }
  $result_out->data =  $resultArray;
  $result_out->status = (true);
  
}else{
  $result_out->data =  $resultArray;
  $result_out->status = (false);
}

  $SearchArray = [

  ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/checkcpart', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
    $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
    
   
   $_sql = "SELECT DISTINCT CPART_NO,USAGE CPART_NAME FROM OE_FABRIC_BILL_HIS
     WHERE SO_YEAR>='".$so_year."' AND SO_NO='".$so_no."'
     ORDER BY CPART_NO";
    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
   
    
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/checksearch', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
    $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
  
    
   
    $_sql = "SELECT * FROM OE_MU_REPORT_V where SO_NO_DOC like '%".$so_no_doc."%' order by MU_SEQ desc";
    

    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
    
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
    
    
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/check_history_per', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
  
    
   
    $_sql = "select * from OE_CUT_MU_HISTORY_PER_M WHERE ORG = '".$org."' and  year >= EXTRACT (year from SYSDATE)-4 ORDER BY YEAR ASC";
    

    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
    
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
    
    
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/checkgmno', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

 
 
    $cpart_no = isset($_REQUEST['CPART']) ? $_REQUEST['CPART'] : '';
    $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
    
    
   
   $_sql = "SELECT * FROM OE_CUT_MU_DETAIL WHERE MU_SEQ = '".$mu_seq."'";
    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });
  
  $app->post('/export_data', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    

    
    
  $_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD') and ORG = '".$org."'
  order by MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/export_data_2', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    $f_type = isset($_REQUEST['f_type']) ? $_REQUEST['f_type'] : '';
    

    
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE)  =  TO_DATE('".$start."','RRRR/MM/DD') 
  and ORG = '".$org."' and F_TYPE = '".$f_type."'
  order by MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/export_data_3', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    $f_type = isset($_REQUEST['f_type']) ? $_REQUEST['f_type'] : '';
    

    
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD') 
  and ORG = '".$org."' and F_TYPE = '".$f_type."'
  order by MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/export_data_4', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';


    
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD') 
  and ORG = '".$org."'
  order by WITH_LOSS ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  

  $app->post('/export_data_5', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';


    
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')
  and ORG = '".$org."'
  order by MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/export_data_weekly', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

    

    
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')
  and ORG = '".$org."' order by MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/export_data_weekly_analy', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

    

    
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')
  and ORG = '".$org."' order by GM_TABLE ASC , MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/export_data_all_weekly', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['first_day']) ? $_REQUEST['first_day'] : '';
    $end = isset($_REQUEST['last_day']) ? $_REQUEST['last_day'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

    

    
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')
  and ORG = '".$org."' order by MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
    $result_out->start = $start;
    $result_out->end = $end;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
    $result_out->start = $start;
    $result_out->end = $end;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/check_data_have', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $multiple = isset($_REQUEST['multiple']) ? $_REQUEST['multiple'] : '';
    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

    
   
    
$_sql = "SELECT distinct (f_type) f_type FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$start."','RRRR/MM/DD')
and ORG = '".$org."' and f_type = '".$multiple."'";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;

  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;

  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/check_data_have_weekly', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $multiple = isset($_REQUEST['multiple']) ? $_REQUEST['multiple'] : '';
    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

    
   
    
$_sql = "SELECT distinct (f_type) f_type FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')
and ORG = '".$org."' and f_type = '".$multiple."'";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;

  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;

  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/export_data_all_type', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

  
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$start."','RRRR/MM/DD')
  and ORG = '".$org."' order by MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/export_data_all_weekly_3_1', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

    

    
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')
  and ORG = '".$org."' order by MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/export_data_all_type_3_1', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

  
    
$_sql = "SELECT * FROM OE_MU_REPORT_V WHERE TRUNC(MU_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')
  and ORG = '".$org."' order by MU_DATE ASC";
  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    foreach ($DataRows as $row){
      array_push($resultArray,$row);
    }
    $result_out->data =  $resultArray;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/export_data_all_weekly_cut', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

    

  /*   select SUM(CUT_QTY) CUT_QTY from nyg1.dfit_cut_detail_all_bu@wcs.world
where BU_CODE in('G2') and TRUNC(cut_date) =  TO_DATE('05/05/2022','DD/MM/RRRR');
     */
$_sql = "SELECT SUM(CUT_QTY) CUT_QTY FROM nyg1.dfit_cut_detail_all_bu@wcs.world where
TRUNC(CUT_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')  and  BU_CODE = '".$org."'";

  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
  
    $result_out->data =  $DataRows;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $DataRows;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/export_data_all_weekly_yard', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

    
    
  
    $_sql = "SELECT SUM(FABRIC_WEIGHT) FABRIC_WEIGHT FROM nyg1.dfit_cut_head_all_bu@wcs.world where
    TRUNC(CUT_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')  and  BU_CODE = '".$org."'";

  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    
    $result_out->data =  $DataRows;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $DataRows;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/export_data_all_weekly_yard_2', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
    $end = isset($_REQUEST['end']) ? $_REQUEST['end'] : '';
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

    
    
  
    $_sql = "SELECT SUM(FABRIC_WEIGHT) FABRIC_WEIGHT FROM nyg1.dfit_cut_head_p_all_bu@wcs.world where
    TRUNC(CUT_DATE) between TO_DATE('".$start."','RRRR/MM/DD') and TO_DATE('".$end."','RRRR/MM/DD')  and  BU_CODE = '".$org."'";

  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    
    $result_out->data =  $DataRows;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $DataRows;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->get('/get_target', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


    $_sql = "SELECT * FROM OE_CUT_MU_TARGET";
    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if ($DataRows != false) {
      foreach ($DataRows as $row) {
        array_push($resultArray, $row);
      }
      $result_out->data =  $resultArray;
      $result_out->status = (true);
    } else {
      $result_out->data =  $resultArray;
      $result_out->status = (false);
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->get('/get_date', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


    $_sql = "SELECT TO_CHAR(START_DATE,'DD/MM/RRRR')START_DATE  FROM OE_CUT_MU_DATE_PROCESS";
    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if ($DataRows != false) {
      foreach ($DataRows as $row) {
        array_push($resultArray, $row);
      }
      $result_out->data =  $resultArray;
      $result_out->status = (true);
    } else {
      $result_out->data =  $resultArray;
      $result_out->status = (false);
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->get('/get_date2', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


    $_sql = "SELECT TO_CHAR(START_DATE,'RRRR/MM/DD')START_DATE  FROM OE_CUT_MU_DATE_PROCESS";
    $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if ($DataRows != false) {
      foreach ($DataRows as $row) {
        array_push($resultArray, $row);
      }
      $result_out->data =  $resultArray;
      $result_out->status = (true);
    } else {
      $result_out->data =  $resultArray;
      $result_out->status = (false);
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/update_target', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $end_loss_woven = isset($_REQUEST['end_woven']) ? $_REQUEST['end_woven'] : '';;
    $end_loss_light = isset($_REQUEST['end_light']) ? $_REQUEST['end_light'] : '';;
    $end_loss_fleece = isset($_REQUEST['end_fleece']) ? $_REQUEST['end_fleece'] : '';;
    $end_loss_special = isset($_REQUEST['end_special']) ? $_REQUEST['end_special'] : '';;

    $splice_loss_woven = isset($_REQUEST['splice_woven']) ? $_REQUEST['splice_woven'] : '';;
    $splice_loss_light = isset($_REQUEST['splice_light']) ? $_REQUEST['splice_light'] : '';;
    $splice_loss_fleece = isset($_REQUEST['splice_fleece']) ? $_REQUEST['splice_fleece'] : '';;
    $splice_loss_special = isset($_REQUEST['splice_special']) ? $_REQUEST['splice_special'] : '';;

    $cut_loss_woven = isset($_REQUEST['cut_woven']) ? $_REQUEST['cut_woven'] : '';;
    $cut_loss_light = isset($_REQUEST['cut_light']) ? $_REQUEST['cut_light'] : '';;
    $cut_loss_fleece = isset($_REQUEST['cut_fleece']) ? $_REQUEST['cut_fleece'] : '';;
    $cut_loss_special = isset($_REQUEST['cut_special']) ? $_REQUEST['cut_special'] : '';;

    $rem_loss_woven = isset($_REQUEST['rem_woven']) ? $_REQUEST['rem_woven'] : '';;
    $rem_loss_light = isset($_REQUEST['rem_light']) ? $_REQUEST['rem_light'] : '';;
    $rem_loss_fleece = isset($_REQUEST['rem_fleece']) ? $_REQUEST['rem_fleece'] : '';;
    $rem_loss_special = isset($_REQUEST['rem_special']) ? $_REQUEST['rem_special'] : '';;

    $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';;

    
  
  
    $_sql = "UPDATE OE_CUT_MU_TARGET SET END_LOSS_WOVEN = '" . $end_loss_woven . "',END_LOSS_LIGHT = '" . $end_loss_light . "'
    ,END_LOSS_FLEECE =  '" . $end_loss_fleece . "',END_LOSS_SPECIAL = '" . $end_loss_special. "',SPLICE_LOSS_WOVEN = '" . $splice_loss_woven . "',SPLICE_LOSS_LIGHT = '" . $splice_loss_light . "'
    ,SPLICE_LOSS_FLEECE = '" . $splice_loss_fleece . "',SPLICE_LOSS_SPECIAL = '" . $splice_loss_special . "',CUT_LOSS_WOVEN = '" . $cut_loss_woven . "'
    ,CUT_LOSS_LIGHT = '" . $cut_loss_light . "', CUT_LOSS_FLEECE = '".$cut_loss_fleece."' ,CUT_LOSS_SPECIAL = '" . $cut_loss_special . "'
    ,REM_LOSS_WOVEN = '" . $rem_loss_woven . "',REM_LOSS_LIGHT ='" . $rem_loss_light . "',REM_LOSS_FLEECE = '" . $rem_loss_fleece . "',REM_LOSS_SPECIAL = '" . $rem_loss_special . "'
    , USER_UPDATE ='".$username."',UPDATE_DATE = SYSDATE WHERE SEQ = '1'";
    
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->sql = ($_sql);
    } else {
  
      $result_out->status = (false);
      $result_out->sql = ($_sql);
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/update_start_date', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    
    $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';;
    $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';;

    
  
  
    $_sql = "UPDATE OE_CUT_MU_DATE_PROCESS SET START_DATE = TO_DATE('" . $start . "','DD/MM/RRRR')
    , UPDATE_BY ='".$username."',UPDATE_DATE = SYSDATE WHERE SEQ = '1'";
  
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->sql = ($_sql);
    } else {
  
      $result_out->status = (false);
      $result_out->sql = ($_sql);
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/update_history_per', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $year = isset($_REQUEST['year']) ? $_REQUEST['year'] : '';;
    $end_loss = isset($_REQUEST['end_loss']) ? $_REQUEST['end_loss'] : '';;
    $splice_loss = isset($_REQUEST['splice_loss']) ? $_REQUEST['splice_loss'] : '';;
    $cut_out_loss = isset($_REQUEST['cut_out_loss']) ? $_REQUEST['cut_out_loss'] : '';;
    $remnants_loss = isset($_REQUEST['remnants_loss']) ? $_REQUEST['remnants_loss'] : '';;
    $total_spread_loss_per = isset($_REQUEST['total_spread_loss_per']) ? $_REQUEST['total_spread_loss_per'] : '';;
    $total_width_loss_per = isset($_REQUEST['total_width_loss_per']) ? $_REQUEST['total_width_loss_per'] : '';;
    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';;
    $index_no  = isset($_REQUEST['index_no']) ? $_REQUEST['index_no'] : '';;


  
    $_sql = "UPDATE OE_CUT_MU_HISTORY_PER_M SET YEAR = '".$year."' , END_LOSS = '".$end_loss."' , SPLICE_LOSS = '".$splice_loss."'
    , CUT_OUT_LOSS = '".$cut_out_loss."' , REMNANTS_LOSS = '".$remnants_loss."' , TOTAL_SPREAD_LOSS_PER = '".$total_spread_loss_per."'
    , TOTAL_WIDTH_LOSS_PER = '".$total_width_loss_per."' WHERE ORG = '".$org."' and INDEX_NO = '".$index_no."'";
   
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->sql = ($_sql);
    } else {
  
      $result_out->status = (false);
      $result_out->sql = ($_sql);
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

 

  $app->post('/find_history_data', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';

  
    $_sql = "SELECT * FROM OE_CUT_MU_HISTORY_PER_M WHERE ORG = '".$org."'";

  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    
    $result_out->data =  $DataRows;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $DataRows;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/check_year_duplicate', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    $year = isset($_REQUEST['year']) ? $_REQUEST['year'] : '';


  
    $_sql = "SELECT * FROM OE_CUT_MU_HISTORY_PER_M WHERE ORG = '".$org."' and YEAR = '".$year."'";

  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    
    $result_out->data =  $DataRows;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $DataRows;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/check_max_value', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    $year = isset($_REQUEST['year']) ? $_REQUEST['year'] : '';


  
    $_sql = "SELECT MAX(INDEX_NO) LAST_INDEX FROM OE_CUT_MU_HISTORY_PER_M WHERE ORG = '".$org."'";

  
 $result_out = new stdClass();
    $DataRows = [];
    $resultArray = array(); //data
    $DataRows = ConnectDbAll($_sql);
  
    if($DataRows != false){
    
    $result_out->data =  $DataRows;
    $result_out->status = (true);
    $result_out->sql = $_sql;
  }else{
    $result_out->data =  $DataRows;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }
  
    $SearchArray = [
  
    ];
  
  $resultArray2 = array();
  array_push($resultArray2,$SearchArray);  
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/insert_new_value_year', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


    $org = isset($_REQUEST['org']) ? $_REQUEST['org'] : '';
    $year = isset($_REQUEST['year']) ? $_REQUEST['year'] : '';
    $end_loss = isset($_REQUEST['end_loss']) ? $_REQUEST['end_loss'] : '';
    $splice_loss = isset($_REQUEST['splice_loss']) ? $_REQUEST['splice_loss'] : '';
    $cut_out_loss = isset($_REQUEST['cut_out_loss']) ? $_REQUEST['cut_out_loss'] : '';
    $remnants_loss = isset($_REQUEST['remnants_loss']) ? $_REQUEST['remnants_loss'] : '';
    $total_spread_loss = isset($_REQUEST['total_spread_loss']) ? $_REQUEST['total_spread_loss'] : '';
    $total_width_loss = isset($_REQUEST['total_width_loss']) ? $_REQUEST['total_width_loss'] : '';
    $index_no = isset($_REQUEST['max_value']) ? $_REQUEST['max_value'] : '';



  
    $_sql = "INSERT INTO OE_CUT_MU_HISTORY_PER_M (YEAR , END_LOSS , SPLICE_LOSS , CUT_OUT_LOSS , REMNANTS_LOSS , TOTAL_SPREAD_LOSS_PER , 
    TOTAL_WIDTH_LOSS_PER , ORG , INDEX_NO) VALUES ('".$year."' , '".$end_loss."' , '".$splice_loss."' , '".$cut_out_loss."' , '".$remnants_loss."'
    , '".$total_spread_loss."' , '".$total_width_loss."' , '".$org."' , '".$index_no."')";

  
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->sql = ($_sql);
    } else {
  
      $result_out->status = (false);
      $result_out->sql = ($_sql);
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
  });


  

$app->run(); //สั่งระบบให้ทำงาน

?>
