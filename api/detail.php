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
    oci_commit($conn);
    if($objEx){
      
    }else{
      echo "Connect Data Base error";
    }
  }
  oci_close($conn);
  return $DataRows;

}

 $app->post('/checkso', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

 
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';

  
  
 
  $_sql = "SELECT MU_SEQ, TO_CHAR(MU_DATE,'DD/MM/RRRR')MU_DATE, SO_YEAR , SO_NO , SO_NO_DOC , STYLE_REF , CUSTOMER_NAME , ORDER_QTY , 
  GMT_TYPE , CPART_NO , F_TYPE , FABRIC_TYPE , ORG   FROM OE_CUT_MU_HEAD where SO_NO = '".$so_no."' and SO_YEAR = '".$so_year."' and MU_SEQ 
  = '".$mu_seq."'";
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
$app->post('/findrowgm', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

 

 // $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $gm_no = "210001-02";
  
 
   $_sql = "SELECT * FROM OE_CUT_MU_DETAIL WHERE GM_NO = '".$gm_no."'";
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
$app->post('/updatedata', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
   $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
   $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
   $cut_table = isset($_REQUEST['CUT_TABLE']) ? $_REQUEST['CUT_TABLE'] : '';
   $weight_m2 = isset($_REQUEST['WEIGHT_M2']) ? $_REQUEST['WEIGHT_M2'] : '';
   $weight_yd = isset($_REQUEST['WEIGHT_YD']) ? $_REQUEST['WEIGHT_YD'] : '';
   $width_inch = isset($_REQUEST['WIDTH_INCH']) ? $_REQUEST['WIDTH_INCH'] : '';
   $length_yd = isset($_REQUEST['LENGTH_YD']) ? $_REQUEST['LENGTH_YD'] : '';
   $length_inch= isset($_REQUEST['LENGTH_INCH']) ? $_REQUEST['LENGTH_INCH'] : '';
   $marker = isset($_REQUEST['MARKER']) ? $_REQUEST['MARKER'] : '';
   $ply = isset($_REQUEST['PLY']) ? $_REQUEST['PLY'] : '';

   
  
$_sql = "UPDATE  OE_CUT_MU_DETAIL SET CUT_TABLE = '".$cut_table."',WEIGHT_M2= '".$weight_m2."', WEIGHT_YD ='".$weight_yd."' ,
    WIDTH_INCH='".$width_inch."'  ,LENGTH_YD='".$length_yd."' , LENGTH_INCH='".$length_inch."' ,MARKER ='".$marker."', PLY='".$ply."' 
    WHERE GM_NO = '".$gm_no."' and MU_SEQ = '".$mu_seq."'";
   $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];
 
 $resultArray2 = array();
 array_push($resultArray2,$SearchArray);
 $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
 
   return $response; // ส่งคำตอบกลับ
 });
 $app->post('/getdatamain', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

 
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  
  
 
   $_sql = "SELECT *  FROM OE_CUT_MU_HEAD WHERE SO_NO = '".$so_no."' and SO_YEAR =  '".$so_year."' and MU_SEQ = '".$mu_seq."'";
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

$app->post('/checkgmno', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

 
 
  $cpart_no = isset($_REQUEST['CPART']) ? $_REQUEST['CPART'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  
  
 
 $_sql = "SELECT * FROM OE_CUT_MU_DETAIL WHERE MU_SEQ = '".$mu_seq."' ORDER BY CREATE_DATE ASC";
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
$app->post('/show_sub_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
  
  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
 
  $_sql = "SELECT * FROM OE_CUT_MU_DETAIL_SUB where GM_NO = '".$gm_no."' and MU_SEQ = '".$mu_seq."'";
 
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




$app->post('/insert_sub_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
  $gm_no = isset($_REQUEST['GM_NO2']) ? $_REQUEST['GM_NO2'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ2']) ? $_REQUEST['MU_SEQ2'] : '';
  $issue = isset($_REQUEST['ISSUE2']) ? $_REQUEST['ISSUE2'] : '';
  $issue_yard = isset($_REQUEST['ISSUE2_YARD']) ? $_REQUEST['ISSUE2_YARD'] : '';
  $avg_width = isset($_REQUEST['AVG_WIDTH2']) ? $_REQUEST['AVG_WIDTH2'] : '';
  $line_no = isset($_REQUEST['LINE_NO']) ? $_REQUEST['LINE_NO'] : '';
  
 
  $_sql = "INSERT INTO OE_CUT_MU_DETAIL_SUB (GM_NO,MU_SEQ,ISSUE,ISSUE_YRD,AVG_WIDTH,LINE_NO) VALUES ('".$gm_no."','".$mu_seq."'
 ,'".$issue."','".$issue_yard."','".$avg_width."','".$line_no."')";

 $_sql2 = "COMMIT;";
 
  $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql,$_sql2);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});


$app->post('/find_line_no', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  /* $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : ''; */
  $_sql = "SELECT MU_LINE_SEQ.NEXTVAL FROM DUAL";
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
 $app->post('/update_sub_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url
  
  $gm_no = isset($_REQUEST['GM_NO2']) ? $_REQUEST['GM_NO2'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ2']) ? $_REQUEST['MU_SEQ2'] : '';
  $issue_yard = isset($_REQUEST['ISSUE2_YARD']) ? $_REQUEST['ISSUE2_YARD'] : '';
  $issue = isset($_REQUEST['ISSUE2']) ? $_REQUEST['ISSUE2'] : '';
  $avg_width = isset($_REQUEST['AVG_WIDTH2']) ? $_REQUEST['AVG_WIDTH2'] : '';
  $line_no = isset($_REQUEST['LINE_NO']) ? $_REQUEST['LINE_NO'] : '';
 
  $_sql = "UPDATE OE_CUT_MU_DETAIL_SUB SET ISSUE = '".$issue."',ISSUE_YRD = '".$issue_yard."' ,AVG_WIDTH = '".$avg_width."' WHERE MU_SEQ = '".$mu_seq."' and  GM_NO = '".$gm_no."' and LINE_NO = '".$line_no."'";
 
  $result_out = new stdClass();
 
  $DataRows = ConnectDbnoAll($_sql);

  if($DataRows != false){
  
 
  $result_out->status = (true);
  $result_out->sql = ($_sql);
  
}else{
  
  $result_out->status = (false);
  $result_out->sql = ($_sql);
}

  $SearchArray = [

  ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
 
   return $response; // ส่งคำตอบกลับร้า
 });

 $app->post('/delete_sub_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  $line_no = isset($_REQUEST['LINE_NO']) ? $_REQUEST['LINE_NO'] : '';
  
 
  $_sql = "DELETE FROM OE_CUT_MU_DETAIL_SUB WHERE GM_NO = '".$gm_no."' and MU_SEQ = '".$mu_seq."' and LINE_NO ='".$line_no."'";

 $_sql2 = "COMMIT;";
 
  $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql,$_sql2);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/check_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $cpart_no = isset($_REQUEST['CPART']) ? $_REQUEST['CPART'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  
 
$_sql = "SELECT DISTINCT  GM_NO,GET_MU_GM2(SO_YEAR,SO_NO,OU_CODE,CPART_NO) GM_M2,G_Y GM_YD,WIDTH WIDTH_INCH,ROUND(MARK_YARD+MARK_LENGTH/36,2)  YD,MARK_YARD LENGTH_YD,MARK_LENGTH LENGTH_INCH,PER_U,LAYER 
 FROM OE_CUT_PLAN_GM08 G
 WHERE SO_YEAR='".$so_year."' AND SO_NO='".$so_no."' AND CPART_NO='".$cpart_no."' AND GM_NO ='".$gm_no."'
 AND GM_NO NOT IN (SELECT D.GM_NO FROM OE_CUT_MU_DETAIL D,OE_CUT_MU_HEAD H
 WHERE H.MU_SEQ=D.MU_SEQ AND MU_STATUS='OPEN' 
 AND G.GM_NO=D.GM_NO)
 ORDER  BY GM_NO"; 
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


$app->post('/all_data_bottom', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
  $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  
  
  $_sql = "SELECT * FROM OE_MU_REPORT_V WHERE SO_NO_DOC = '".$so_no_doc."' and MU_SEQ = '".$mu_seq."' "; 
  
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
  $result_out->sql = ($_sql);
  $result_out->datarow = $DataRows;
  
}else{
  $result_out->data =  $resultArray;
  $result_out->status = (false);
  $result_out->sql = ($_sql);
}


  $SearchArray = [

  ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});


$app->post('/update_bottom', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
  $end_plug1 = isset($_REQUEST['end_plug1']) ? $_REQUEST['end_plug1'] : '';
  $end_plug2 = isset($_REQUEST['end_plug2']) ? $_REQUEST['end_plug2'] : '';
  $casue = isset($_REQUEST['casue']) ? $_REQUEST['casue'] : '';
  $code = isset($_REQUEST['code']) ? $_REQUEST['code'] : '';
  $endloss_auto = isset($_REQUEST['endloss_auto']) ? $_REQUEST['endloss_auto'] : '';
  $endloss_inch = isset($_REQUEST['endloss_inch']) ? $_REQUEST['endloss_inch'] : '';
  $splicenoof = isset($_REQUEST['splicenoof']) ? $_REQUEST['splicenoof'] : '';
  $spliceinch = isset($_REQUEST['spliceinch']) ? $_REQUEST['spliceinch'] : '';
  $cutoutnoof = isset($_REQUEST['cutoutnoof']) ? $_REQUEST['cutoutnoof'] : '';
  $cutoutweight = isset($_REQUEST['cutoutweight']) ? $_REQUEST['cutoutweight'] : '';
  $remnants = isset($_REQUEST['remnants']) ? $_REQUEST['remnants'] : '';
  $widthplug1 = isset($_REQUEST['widthplug1']) ? $_REQUEST['widthplug1'] : '';
  $widthplug2 = isset($_REQUEST['widthplug2']) ? $_REQUEST['widthplug2'] : '';
  $avgglueside = isset($_REQUEST['avgglueside']) ? $_REQUEST['avgglueside'] : '';
  $re_cut = isset($_REQUEST['re_cut']) ? $_REQUEST['re_cut'] : '';
  $end_plug1_2_yd = isset($_REQUEST['endplug12_yd']) ? $_REQUEST['endplug12_yd'] : '';
  $plug12_inch = isset($_REQUEST['plug12_inch']) ? $_REQUEST['plug12_inch'] : '';
  $plug12 = isset($_REQUEST['plug12']) ? $_REQUEST['plug12'] : '';
  
 
 $_sql = "UPDATE OE_CUT_MU_HEAD SET END_PLUG1 = '".$end_plug1."',
  END_PLUG2 = '".$end_plug2."',
  CAUSE_CODE = '".$code."',
  CAUSE_NAME = '".$casue."',
  SPLICE = '".$splicenoof."',
  SPLICE_INCH = '".$spliceinch."',
  CUT_OUT = '".$cutoutnoof."',
  CUT_OUT_GRAM = '".$cutoutweight."',
  REMNANTS_WEIGHT = '".$remnants."',
  PLUG1_GRAM = '".$widthplug1."',
  PLUG2_GRAM = '".$widthplug2."',
  END_LOSS_AUTO = '".$endloss_auto."',
  END_LOSS_INCH = '".$endloss_inch."',
  GLUE_SIDE = '".$avgglueside."',
  END_PLUG1_2_YD = '".$end_plug1_2_yd."',
  PLUG1_2_INCH = '".$plug12."',
  RE_CUT = '".$re_cut."'
  WHERE MU_SEQ = '".$mu_seq."' and SO_NO = '".$so_no."' and SO_NO_DOC = '".$so_no_doc."'";

 $_sql2 = "COMMIT;";
 
  $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql,$_sql2);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/delete_head', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
  $so_no_doc = isset($_REQUEST['so_no_doc']) ? $_REQUEST['so_no_doc'] : '';

  
 
 $_sql = "DELETE FROM OE_CUT_MU_HEAD WHERE MU_SEQ = '".$mu_seq."' and SO_NO_DOC = '".$so_no_doc."'";

 $_sql2 = "COMMIT;";
 
  $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql,$_sql2);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/delete_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
  $so_no_doc = isset($_REQUEST['so_no_doc']) ? $_REQUEST['so_no_doc'] : '';

  
 
 $_sql = "DELETE FROM OE_CUT_MU_DETAIL WHERE MU_SEQ = '".$mu_seq."'";

 $_sql2 = "COMMIT;";
 
  $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql,$_sql2);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});


$app->post('/check_endless_auto', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  
 
$_sql = "SELECT * FROM OE_CUT_MU_HEAD where MU_SEQ = '".$mu_seq."' and SO_YEAR = '".$so_year."' and SO_NO='".$so_no."'"; 
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


$app->post('/delete_glue_side', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
 

  
 
 $_sql = "DELETE FROM OE_CUT_MU_GLUE WHERE MU_SEQ = '".$mu_seq."'";

 $_sql2 = "COMMIT;";
 
  $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql,$_sql2);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/delete_splice', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
 

  
 
 $_sql = "DELETE FROM OE_CUT_MU_SPLICE WHERE MU_SEQ = '".$mu_seq."'";

 $_sql2 = "COMMIT;";
 
  $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql,$_sql2);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/delete_detail_sub', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
 

  
 
 $_sql = "DELETE FROM OE_CUT_MU_DETAIL_SUB WHERE MU_SEQ = '".$mu_seq."'";

 $_sql2 = "COMMIT;";
 
  $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql,$_sql2);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/open_avg', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';

  $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';

  
 
$_sql = "SELECT * FROM OE_CUT_MU_HEAD where MU_SEQ = '".$mu_seq."' and SO_NO_DOC ='".$so_no_doc."'"; 
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

$app->post('/create_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  $table = isset($_REQUEST['CUT_TABLE']) ? $_REQUEST['CUT_TABLE'] : '';
  $weight_m2= isset($_REQUEST['WEIGHT_M2']) ? $_REQUEST['WEIGHT_M2'] : '';
  $weight_yd = isset($_REQUEST['WEIGHT_YD']) ? $_REQUEST['WEIGHT_YD'] : '';
  $width_inch = isset($_REQUEST['WIDTH_INCH']) ? $_REQUEST['WIDTH_INCH'] : '';
  $length_yd = isset($_REQUEST['LENGTH_YD']) ? $_REQUEST['LENGTH_YD'] : '';
  $length_inch = isset($_REQUEST['LENGTH_INCH']) ? $_REQUEST['LENGTH_INCH'] : '';
  $marker = isset($_REQUEST['MARKER']) ? $_REQUEST['MARKER'] : '';
  $ply = isset($_REQUEST['PLY']) ? $_REQUEST['PLY'] : '';
  $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
  
 
  $_sql = "INSERT INTO OE_CUT_MU_DETAIL (MU_SEQ,CUT_TABLE,GM_NO,WEIGHT_M2,WEIGHT_YD,WIDTH_INCH,LENGTH_YD,LENGTH_INCH,MARKER,PLY,CREATE_BY,CREATE_DATE)
  VALUES ('".$mu_seq."','".$table."','".$gm_no."','".$weight_m2."','".$weight_yd."','".$width_inch."','".$length_yd."','".$length_inch."'
  ,'".$marker."','".$ply."','".$username."',SYSDATE)";

 $_sql2 = "COMMIT;";
 
  $result_out = new stdClass();
 
   $DataRows = ConnectDbnoAll($_sql,$_sql2);
 
   if($DataRows != false){
   
  
   $result_out->status = (true);
   $result_out->sql = ($_sql);
   
 }else{
   
   $result_out->status = (false);
   $result_out->sql = ($_sql);
 }
 
   $SearchArray = [
 
   ];

$resultArray2 = array();
array_push($resultArray2,$SearchArray);
$response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/check_data_have', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  
  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  
 
$_sql = "SELECT * FROM OE_CUT_MU_DETAIL where MU_SEQ = '".$mu_seq."' and GM_NO = '".$gm_no."'"; 
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



$app->run(); //สั่งระบบให้ทำงาน



?>