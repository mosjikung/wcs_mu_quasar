<?php
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: access");
header("Access-Control-Allow-Methods: GET,POST");
header("Access-Control-Allow-Credentials: true");
header('Content-Type: application/json;charset=utf-8');
require_once __DIR__ . '/../api/web_config.php';

set_time_limit(60000);

use \Psr\Http\Message\ResponseInterface as Response; // ไลบราลี้สำหรับจัดการคำร้องขอ
use \Psr\Http\Message\ServerRequestInterface as Request; // ไลบราลี้สำหรับจัดการคำตอบกลับ

require './vendor/autoload.php'; // ดึงไฟ์ autoload.php เข้ามา
//include_once './class.oracle.php'; // Class Connect Oracle
include_once './util.php'; // ดึงไฟ์ util.php เข้ามา
include_once './web_config.php';
$app = new \Slim\App; // สร้าง object หลักของระบบ

date_default_timezone_set("Asia/Bangkok");

function ConnectDbAll($_sql)
{
  $DataRows = array();
  $conn = ConnectedDBSO();
  if (!$conn) {
    $_err = oci_error();
    echo $_err;
  } else {
    $objParse = oci_parse($conn, $_sql);
    $objEx = oci_execute($objParse);
    if ($objEx) {
      $objResult = oci_fetch_all($objParse, $DataRows, null, null, OCI_FETCHSTATEMENT_BY_ROW);
    } else {
      echo "Connect Data Base error";
    }
  }
  oci_close($conn);
  return $DataRows;
}
function ConnectDbnoAll($_sql)
{
  $DataRows = array();
  $conn = ConnectedDBSO();
  if (!$conn) {
    $_err = oci_error();
    echo $_err;
  } else {
    $objParse = oci_parse($conn, $_sql);
    $objEx = oci_execute($objParse);
    if ($objEx) {
    } else {
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
  (SELECT SUM(RATIO*LAYER) CUT_QTY FROM OE_CUT_PLAN_GM08 where so_no=OE_ME_ORDER.SO_NO and so_year=OE_ME_ORDER.SO_YEAR ) CUT_QTY,
  SHIPMENT_DATE, GMT_TYPE,FABRIC_DESCRIPTION FABRIC_TYPE 
  FROM OE_ME_ORDER WHERE SO_YEAR>='".$so_year."' AND SO_NO='".$so_no."'";
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
    $result_out->sql = $_sql;
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});


$app->post('/checkso2', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $cpart_no = isset($_REQUEST['cpartno']) ? $_REQUEST['cpartno'] : '';


  $_sql = "select DISTINCT ITEM_DESC from oe_so_bom  where so_No='".$so_no."' and so_year='".$so_year."' and cpart_No='".$cpart_no."'";
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
    $result_out->sql = $_sql;
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/check_fabric', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $cpart_no = isset($_REQUEST['cpartno']) ? $_REQUEST['cpartno'] : '';


  $_sql = "select DISTINCT ITEM_DESC from oe_so_bom  where so_No='".$so_no."' and so_year='".$so_year."' and cpart_No='".$cpart_no."'";
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
    $result_out->sql = $_sql;
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/checkcpart', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';


  $_sql = "SELECT DISTINCT CPART_NO,USAGE CPART_NAME FROM OE_FABRIC_BILL_HIS
     WHERE SO_YEAR>='" . $so_year . "' AND SO_NO='" . $so_no . "'
     ORDER BY CPART_NO";
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


$app->post('/checkgm', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $cpart_no = isset($_REQUEST['cpartno']) ? $_REQUEST['cpartno'] : '';


  $_sql = "SELECT DISTINCT GM_NO FROM OE_CUT_PLAN_GM08
  WHERE SO_YEAR='".$so_year."' AND SO_NO='".$so_no."' AND CPART_NO='".$cpart_no."' AND GM_NO NOT IN (SELECT GM_NO FROM OE_CUT_MU_DETAIL)
  ORDER BY GM_NO";
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
    $result_out->sql = ($_sql);
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = ($_sql);
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});
$app->get('/createmu', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $_sql = "SELECT MU_SEQ.NEXTVAL FROM DUAL";
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
$app->get('/create_line_no', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $_sql = "SELECT MU_LINE_SEQ.NEXTVAL FROM DUAL";
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


$app->get('/detail_start_date', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $_sql = "SELECT TO_CHAR(START_DATE,'RRRR/MM/DD')START_DATE , TO_CHAR(END_DATE,'RRRR/MM/DD')END_DATE FROM OE_CUT_MU_DATE_PROCESS";
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
    $result_out->sql = $_sql;
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->get('/detail_start_date_first', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


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
    $result_out->sql = $_sql;
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});




$app->post('/createdetail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
  $GM_NO = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $table = isset($_REQUEST['table']) ? $_REQUEST['table'] : '';
  $weightm2 = isset($_REQUEST['weightm2']) ? $_REQUEST['weightm2'] : '';
  $weightyd = isset($_REQUEST['weightyd']) ? $_REQUEST['weightyd'] : '';
  $widthinch = isset($_REQUEST['widthinch']) ? $_REQUEST['widthinch'] : '';
  $lenthyd = isset($_REQUEST['lenthyd']) ? $_REQUEST['lenthyd'] : '';
  $lenthyds = isset($_REQUEST['lenthyds']) ? $_REQUEST['lenthyds'] : '';
  $lenthinch = isset($_REQUEST['lenthinch']) ? $_REQUEST['lenthinch'] : '';
  $ttlmarker = isset($_REQUEST['ttlmarker']) ? $_REQUEST['ttlmarker'] : '';
  $marker = isset($_REQUEST['marker']) ? $_REQUEST['marker'] : '';
  $ply = isset($_REQUEST['ply']) ? $_REQUEST['ply'] : '';
  $time = date('d/m/Y');
  $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';


  $_sql = "INSERT INTO OE_CUT_MU_DETAIL (MU_SEQ,CUT_TABLE,GM_NO,WEIGHT_M2,WEIGHT_YD,WIDTH_INCH,LENGTH_YD,LENGTH_INCH,MARKER,PLY,CREATE_BY,CREATE_DATE)
   VALUES ('" . $mu_seq . "','" . $table . "','" . $GM_NO . "','" . $weightm2 . "','" . $weightyd . "','" . $widthinch . "','" . $lenthyd . "','" . $lenthinch . "','" . $marker . "',
   '" . $ply . "','" . $username . "',SYSDATE)";
  $result_out = new stdClass();


  $DataRows = ConnectDbnoAll($_sql);

  if ($DataRows !== false) {


    $result_out->status = (true);
    $result_out->MU_SEQ = $mu_seq;
    $result_out->SQL = $_sql;
  } else {

    $result_out->status = (false);
    $result_out->MU_SEQ = $mu_seq;
    $result_out->SQL = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});


$app->post('/createsubdetail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $MU_SEQ = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  $GM_NO = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $ISSUE_YARD = isset($_REQUEST['ISSUE_YARD']) ? $_REQUEST['ISSUE_YARD'] : '';
  $ISSUE = isset($_REQUEST['ISSUE']) ? $_REQUEST['ISSUE'] : '';
  $AVG_WIDTH = isset($_REQUEST['AVG_WIDTH']) ? $_REQUEST['AVG_WIDTH'] : '';
  $LINE_NO = isset($_REQUEST['LINE_NO']) ? $_REQUEST['LINE_NO'] : '';
  $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';


    $_sql = "INSERT INTO OE_CUT_MU_DETAIL_SUB (MU_SEQ,GM_NO,ISSUE,ISSUE_YRD,AVG_WIDTH,CREATE_BY,CREATE_DATE,LINE_NO)
   VALUES ('" . $MU_SEQ . "','" . $GM_NO . "','" . $ISSUE . "','" . $ISSUE_YARD . "','" . $AVG_WIDTH . "','" . $username . "',SYSDATE,'" . $LINE_NO . "')";
  $result_out = new stdClass();


  $DataRows = ConnectDbnoAll($_sql);

  if ($DataRows !== false) {


    $result_out->status = (true);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = $_sql;
  } else {

    $result_out->status = (false);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/createdetailsub', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $MU_SEQ = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  $GM_NO = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $ISSUE = isset($_REQUEST['ISSUE']) ? $_REQUEST['ISSUE'] : '';
  $AVG_WIDTH = isset($_REQUEST['AVG_WIDTH']) ? $_REQUEST['AVG_WIDTH'] : '';
  $LINE_NO = isset($_REQUEST['LINE_NO']) ? $_REQUEST['LINE_NO'] : '';
  $username = isset($_REQUEST['USERNAME']) ? $_REQUEST['USERNAME'] : '';


  $_sql = "INSERT INTO OE_CUT_MU_DETAIL_SUB (MU_SEQ,GM_NO,ISSUE,AVG_WIDTH,CREATE_BY,CREATE_DATE,LINE_NO)
   VALUES ('" . $MU_SEQ . "','" . $GM_NO . "','" . $ISSUE . "','" . $AVG_WIDTH . "','" . $username . "',SYSDATE,'" . $LINE_NO . "')";
  $result_out = new stdClass();


  $DataRows = ConnectDbnoAll($_sql);

  if ($DataRows !== false) {


    $result_out->status = (true);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = ($_sql);
  } else {

    $result_out->status = (false);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = ($_sql);
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/update_head', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $start = isset($_REQUEST['start']) ? $_REQUEST['start'] : '';
  $f_type = isset($_REQUEST['f_type']) ? $_REQUEST['f_type'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
  $end_plug1 = isset($_REQUEST['END_PLUG1']) ? $_REQUEST['END_PLUG1'] : '';
  $end_plug2 = isset($_REQUEST['END_PLUG2']) ? $_REQUEST['END_PLUG2'] : '';
  $cause_code = isset($_REQUEST['CAUSE_CODE']) ? $_REQUEST['CAUSE_CODE'] : '';
  $cause_name = isset($_REQUEST['CAUSE_NAME']) ? $_REQUEST['CAUSE_NAME'] : '';
  $splice = isset($_REQUEST['SPLICE_NOOFF']) ? $_REQUEST['SPLICE_NOOFF'] : '';
  $splice_inch = isset($_REQUEST['SPICE_INCH']) ? $_REQUEST['SPICE_INCH'] : '';
  $cut_out = isset($_REQUEST['CUT_OUT']) ? $_REQUEST['CUT_OUT'] : '';
  $cut_out_gram = isset($_REQUEST['CUT_OUT_GRAM']) ? $_REQUEST['CUT_OUT_GRAM'] : '';
  $remnants_weight = isset($_REQUEST['REMNANTS_WEIGHT']) ? $_REQUEST['REMNANTS_WEIGHT'] : '';
  $plug1_gram = isset($_REQUEST['PLUG1_GRAM']) ? $_REQUEST['PLUG1_GRAM'] : '';
  $plug2_gram = isset($_REQUEST['PLUG2_GRAM']) ? $_REQUEST['PLUG2_GRAM'] : '';
  $glue_side = isset($_REQUEST['GLUE_SIDE']) ? $_REQUEST['GLUE_SIDE'] : '';
  $endloss_auto = isset($_REQUEST['ENDLOSS_AUTO']) ? $_REQUEST['ENDLOSS_AUTO'] : '';
  $endloss_inch = isset($_REQUEST['ENDLOSS_AUTO_INCH']) ? $_REQUEST['ENDLOSS_AUTO_INCH'] : '';
  $end_plug1_2_yd = isset($_REQUEST['END_PLUG1_2_YD']) ? $_REQUEST['END_PLUG1_2_YD'] : '';
  $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
  $mu_doc_no = isset($_REQUEST['MU_DOC_NO']) ? $_REQUEST['MU_DOC_NO'] : '';
  $plug12 = isset($_REQUEST['PLUG12']) ? $_REQUEST['PLUG12'] : '';
  $re_cut = isset($_REQUEST['RE_CUT']) ? $_REQUEST['RE_CUT'] : '';


  $username = 'IT-LOGON';


  $_sql = "UPDATE OE_CUT_MU_HEAD SET MU_DATE = TO_DATE('".$start."','DD/MM//RRRR') , F_TYPE = '".$f_type."' , END_PLUG1 = '" . $end_plug1 . "',END_PLUG2 = '" . $end_plug2 . "',CAUSE_CODE =  '" . $cause_code . "',CAUSE_NAME = '" . $cause_name . "',SPLICE = '" . $splice . "',SPLICE_INCH = '" . $splice_inch . "',CUT_OUT = '" . $cut_out . "',END_LOSS_AUTO = '" . $endloss_auto . "',END_LOSS_INCH = '" . $endloss_inch . "',CUT_OUT_GRAM = '" . $cut_out_gram . "', END_PLUG1_2_YD = '".$end_plug1_2_yd."' ,REMNANTS_WEIGHT = '" . $remnants_weight . "',PLUG1_GRAM = '" . $plug1_gram . "',PLUG2_GRAM ='" . $plug2_gram . "',GLUE_SIDE = '" . $glue_side . "',MU_DOC_NO = '" . $mu_doc_no . "',PLUG1_2_INCH = '".$plug12."', RE_CUT  = '".$re_cut."', UPDATE_BY ='".$username."',UPDATE_DATE = SYSDATE WHERE MU_SEQ = '" . $mu_seq . "' and SO_NO = '" . $so_no . "' and SO_YEAR = '" . $so_year . "' and SO_NO_DOC = '" . $so_no_doc . "'";
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
$app->post('/returndata', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
 


  $_sql = "SELECT * FROM OE_CUT_MU_DETAIL_SUB where GM_NO = '" . $gm_no . "' and MU_SEQ = '".$mu_seq."'";
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
    $result_out->sql = $_sql;
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});
$app->post('/check_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $so_year =   isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $cpart_no = isset($_REQUEST['CPART_NO']) ? $_REQUEST['CPART_NO'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';


   $_sql = "SELECT DISTINCT  GM_NO,GET_MU_GM2(SO_YEAR,SO_NO,OU_CODE,CPART_NO) GM_M2,G_Y GM_YD,WIDTH WIDTH_INCH,ROUND(MARK_YARD+MARK_LENGTH/36,2)  YD,MARK_YARD LENGTH_YD,MARK_LENGTH LENGTH_INCH,PER_U,LAYER 
   FROM OE_CUT_PLAN_GM08 G
   WHERE SO_YEAR='" . $so_year . "' AND SO_NO='" . $so_no . "' AND CPART_NO='" . $cpart_no . "' AND GM_NO ='" . $gm_no . "'
   AND GM_NO NOT IN (SELECT D.GM_NO FROM OE_CUT_MU_DETAIL D,OE_CUT_MU_HEAD H
   WHERE H.MU_SEQ=D.MU_SEQ AND MU_STATUS='OPEN' 
   AND G.GM_NO=D.GM_NO)
   ORDER  BY GM_NO";
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
    $result_out->sql = ($_sql);
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = ($_sql);
  }


  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});


$app->post('/checkduplicate', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';


  $_sql = "SELECT * FROM OE_CUT_MU_DETAIL where GM_NO = '" . $gm_no . "'";
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

$app->post('/create_for_head', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
  $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
  $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
  $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
  $style = isset($_REQUEST['STYLE']) ? $_REQUEST['STYLE'] : '';
  $customer = isset($_REQUEST['CUSTOMER']) ? $_REQUEST['CUSTOMER'] : '';
  $qty = isset($_REQUEST['QTY']) ? $_REQUEST['QTY'] : '';
  $gm_type = isset($_REQUEST['GM_TYPE']) ? $_REQUEST['GM_TYPE'] : '';
  $f_type = isset($_REQUEST['F_TYPE']) ? $_REQUEST['F_TYPE'] : '';
  $fabric_type = isset($_REQUEST['FABRIC_TYPE']) ? $_REQUEST['FABRIC_TYPE'] : '';
  $date = isset($_REQUEST['DATE']) ? $_REQUEST['DATE'] : '';
  $cpart_no = isset($_REQUEST['CPART_NO']) ? $_REQUEST['CPART_NO'] : '';
  $username = isset($_REQUEST['USERNAME']) ? $_REQUEST['USERNAME'] : '';
  $endloss_auto = isset($_REQUEST['ENDLOSS_AUTO']) ? $_REQUEST['ENDLOSS_AUTO'] : '';
  $endloss_auto_inch = isset($_REQUEST['ENDLOSS_AUTO_INCH']) ? $_REQUEST['ENDLOSS_AUTO_INCH'] : '';
  $date_mu_create = isset($_REQUEST['DATE_MU_CREATE']) ? $_REQUEST['DATE_MU_CREATE'] : '';
  $org = isset($_REQUEST['ORGE']) ? $_REQUEST['ORGE'] : '';

  $customer = str_replace('\'', '', $customer);
  $fabric_type = str_replace('\'', '', $fabric_type);



  $_sql = "INSERT INTO OE_CUT_MU_HEAD (MU_SEQ , MU_DATE , SO_YEAR , SO_NO , SO_NO_DOC , STYLE_REF , CUSTOMER_NAME , ORDER_QTY , GMT_TYPE , F_TYPE,
    CPART_NO , FABRIC_TYPE ,ORG ,END_LOSS_AUTO , END_LOSS_INCH , CREATE_BY , CREATE_DATE , DATE_MU_CREATE) VALUES ('" . $mu_seq . "' ,TO_DATE('".$date_mu_create."','RRRR/MM/DD'), '" . $so_year . "' , '" . $so_no . "' , '" . $so_no_doc . "' , 
    '" . $style . "' , '" . $customer . "' , '" . $qty . "' , '" . $gm_type . "' , '".$f_type."' ,'" . $cpart_no . "' , '" . $fabric_type . "' ,'".$org."','".$endloss_auto."', '".$endloss_auto_inch."' ,'".$username."' , SYSDATE ,
    SYSDATE)";
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
$app->post('/show_head_bottom', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';


  $_sql = "SELECT * FROM OE_MU_REPORT_V where MU_SEQ = '".$mu_seq."' and SO_NO_DOC = '".$so_no_doc."'";
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
    $result_out->sql = ($_sql);
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = ($_sql);
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/update_detail_sub', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
  $line_no = isset($_REQUEST['line_no']) ? $_REQUEST['line_no'] : '';
  $issue_yard = isset($_REQUEST['issue_yard']) ? $_REQUEST['issue_yard'] : '';
  $issue_kg = isset($_REQUEST['issue_kg']) ? $_REQUEST['issue_kg'] : '';
  $avg_width = isset($_REQUEST['avg_width']) ? $_REQUEST['avg_width'] : '';
  $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
 



  $_sql = "UPDATE OE_CUT_MU_DETAIL_SUB SET ISSUE = '".$issue_kg."', ISSUE_YRD = '".$issue_yard."',
  AVG_WIDTH = '".$avg_width."',UPDATE_BY = '".$username."',UPDATE_DATE = SYSDATE where MU_SEQ = '".$mu_seq."' and LINE_NO = '".$line_no."'";

  $_sql2 = "COMMIT";
  $result_out = new stdClass();


  $DataRows = ConnectDbnoAll($_sql,$_sql2);

  if ($DataRows !== false) {


    $result_out->status = (true);
    $result_out->MU_SEQ = $mu_seq;
    $result_out->SQL = $_sql;
  } else {

    $result_out->status = (false);
    $result_out->MU_SEQ = $mu_seq;
    $result_out->SQL = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/delete_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
  $gm_no = isset($_REQUEST['gm_no']) ? $_REQUEST['gm_no'] : '';

 
;


  $_sql = "DELETE FROM OE_CUT_MU_DETAIL where MU_SEQ = '".$mu_seq."' and GM_NO = '".$gm_no."'";

  $_sql2 = "COMMIT";
  $result_out = new stdClass();


  $DataRows = ConnectDbnoAll($_sql,$_sql2);

  if ($DataRows !== false) {


    $result_out->status = (true);
    $result_out->MU_SEQ = $mu_seq;
    $result_out->SQL = $_sql;
  } else {

    $result_out->status = (false);
    $result_out->MU_SEQ = $mu_seq;
    $result_out->SQL = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/delete_detail_sub', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
  $gm_no = isset($_REQUEST['gm_no']) ? $_REQUEST['gm_no'] : '';

 
;


  $_sql = "DELETE FROM OE_CUT_MU_DETAIL_SUB where MU_SEQ = '".$mu_seq."' and GM_NO = '".$gm_no."'";

  $_sql2 = "COMMIT";
  $result_out = new stdClass();


  $DataRows = ConnectDbnoAll($_sql,$_sql2);

  if ($DataRows !== false) {


    $result_out->status = (true);
    $result_out->MU_SEQ = $mu_seq;
    $result_out->SQL = $_sql;
  } else {

    $result_out->status = (false);
    $result_out->MU_SEQ = $mu_seq;
    $result_out->SQL = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/save_data_glue_side', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $MU_SEQ = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  $glue_side1 = isset($_REQUEST['GLUE_SIDE1']) ? $_REQUEST['GLUE_SIDE1'] : '';
  $glue_side2 = isset($_REQUEST['GLUE_SIDE2']) ? $_REQUEST['GLUE_SIDE2'] : '';
  $glue_no = isset($_REQUEST['glue_no']) ? $_REQUEST['glue_no'] : '';

  


  $_sql = "INSERT INTO OE_CUT_MU_GLUE (MU_SEQ,GLUE_SIDE1,GLUE_SIDE2,GLUE_NO)
   VALUES ('" . $MU_SEQ . "','" . $glue_side1 . "','" . $glue_side2. "','" . $glue_no. "')";

  $_sql2 = "COMMIT";
  $result_out = new stdClass();


  $DataRows = ConnectDbnoAll($_sql,$_sql2);

  if ($DataRows !== false) {


    $result_out->status = (true);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = $_sql;
  } else {

    $result_out->status = (false);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->post('/keep_data_for_ply', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  
  $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';


  $_sql = "SELECT * FROM OE_CUT_MU_DETAIL where MU_SEQ = '" . $mu_seq . "'";
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




$app->post('/update_data_glue_side', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $MU_SEQ = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  $glue_side1 = isset($_REQUEST['GLUE_SIDE1']) ? $_REQUEST['GLUE_SIDE1'] : '';
  $glue_side2 = isset($_REQUEST['GLUE_SIDE2']) ? $_REQUEST['GLUE_SIDE2'] : '';
  $glue_no = isset($_REQUEST['GLUE_NO']) ? $_REQUEST['GLUE_NO'] : '';
  


  $_sql = "UPDATE OE_CUT_MU_GLUE SET GLUE_SIDE1 =  '".$glue_side1."' , GLUE_SIDE2 = '".$glue_side2."' where MU_SEQ = '".$MU_SEQ."'
  and GLUE_NO = '".$glue_no."'";

  $_sql2 = "COMMIT";
  $result_out = new stdClass();


  $DataRows = ConnectDbnoAll($_sql,$_sql2);

  if ($DataRows !== false) {


    $result_out->status = (true);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = $_sql;
  } else {

    $result_out->status = (false);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->get('/createmu_glue_side', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $_sql = "SELECT MU_GLUE_SIDE.NEXTVAL FROM DUAL";
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

$app->post('/check_data_glue_side', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

$mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';


  $_sql = "SELECT * FROM OE_CUT_MU_GLUE where MU_SEQ = '".$mu_seq."'";
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
    $result_out->sql = ($_sql);
  } else {
    $result_out->data =  $resultArray;
    $result_out->status = (false);
    $result_out->sql = ($_sql);
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});

$app->get('/create_splice_no', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


  $_sql = "SELECT MU_SPLICE.NEXTVAL FROM DUAL";
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

$app->post('/save_data_splice', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $MU_SEQ = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
  $splice_value = isset($_REQUEST['splice_value']) ? $_REQUEST['splice_value'] : '';
  $splice_no = isset($_REQUEST['splice_no']) ? $_REQUEST['splice_no'] : '';

  


  $_sql = "INSERT INTO OE_CUT_MU_SPLICE (MU_SEQ,SPLICE,SPLICE_NO)
   VALUES ('" . $MU_SEQ . "','" . $splice_value. "','" . $splice_no. "')";

  $_sql2 = "COMMIT";
  $result_out = new stdClass();


  $DataRows = ConnectDbnoAll($_sql,$_sql2);

  if ($DataRows !== false) {


    $result_out->status = (true);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = $_sql;
  } else {

    $result_out->status = (false);
    $result_out->MU_SEQ = $MU_SEQ;
    $result_out->sql = $_sql;
  }

  $SearchArray = [];

  $resultArray2 = array();
  array_push($resultArray2, $SearchArray);
  $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ

  return $response; // ส่งคำตอบกลับร้า
});


$app->post('/check_splice', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

  $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
  
  
    $_sql = "SELECT * FROM OE_CUT_MU_SPLICE where MU_SEQ = '".$mu_seq."'";
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
      $result_out->sql = ($_sql);
    } else {
      $result_out->data =  $resultArray;
      $result_out->status = (false);
      $result_out->sql = ($_sql);
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/update_splice', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $MU_SEQ = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
    $splice_value = isset($_REQUEST['SPLICE_VALUE']) ? $_REQUEST['SPLICE_VALUE'] : '';
    $splice_no = isset($_REQUEST['SPLICE_NO']) ? $_REQUEST['SPLICE_NO'] : '';
    
  
  
    $_sql = "UPDATE OE_CUT_MU_SPLICE SET SPLICE =  '".$splice_value."'  where MU_SEQ = '".$MU_SEQ."'
    and SPLICE_NO = '".$splice_no."'";
  
    $_sql2 = "COMMIT";
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql,$_sql2);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->MU_SEQ = $MU_SEQ;
      $result_out->sql = $_sql;
    } else {
  
      $result_out->status = (false);
      $result_out->MU_SEQ = $MU_SEQ;
      $result_out->sql = $_sql;
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/check_qty_cut', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


    $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
    $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
    $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
    

    $_sql = "SELECT  GM_NO,SO_YEAR,SO_NO,ORDER_QTY,SUM(RATIO*LAYER) CUT_QTY FROM OE_CUT_PLAN_GM08 
    where so_no='".$so_no."'and so_year='".$so_year."' and gm_no ='".$gm_no."'
    GROUP BY  GM_NO,SO_YEAR,SO_NO,ORDER_QTY";
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
      $result_out->sql = $_sql;
    } else {
      $result_out->data =  $resultArray;
      $result_out->status = (false);
      $result_out->sql = $_sql;
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/check_qty_cut_2', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url


    $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
    $so_year = isset($_REQUEST['SO_YEAR']) ? $_REQUEST['SO_YEAR'] : '';
    $so_no = isset($_REQUEST['SO_NO']) ? $_REQUEST['SO_NO'] : '';
    $layer = isset($_REQUEST['PLY']) ? $_REQUEST['PLY'] : '';
    

    $_sql = "SELECT  GM_NO,SO_YEAR,SO_NO,ORDER_QTY,SUM(RATIO*$layer) CUT_QTY FROM OE_CUT_PLAN_GM08 
    where so_no='".$so_no."'and so_year='".$so_year."' and gm_no ='".$gm_no."'
    GROUP BY  GM_NO,SO_YEAR,SO_NO,ORDER_QTY";
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
      $result_out->sql = $_sql;
    } else {
      $result_out->data =  $resultArray;
      $result_out->status = (false);
      $result_out->sql = $_sql;
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/update_detail', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $mu_seq = isset($_REQUEST['mu_seq']) ? $_REQUEST['mu_seq'] : '';
    $gm_no = isset($_REQUEST['GM_NO']) ? $_REQUEST['GM_NO'] : '';
    $table = isset($_REQUEST['table']) ? $_REQUEST['table'] : '';
    $weight_m2 = isset($_REQUEST['weightm2']) ? $_REQUEST['weightm2'] : '';
    $weight_yd = isset($_REQUEST['weightyd']) ? $_REQUEST['weightyd'] : '';
    $width_inch = isset($_REQUEST['widthinch']) ? $_REQUEST['widthinch'] : '';
    $length_yd = isset($_REQUEST['lenthyd']) ? $_REQUEST['lenthyd'] : '';
    $length_inch = isset($_REQUEST['lenthinch']) ? $_REQUEST['lenthinch'] : '';
    $marker = isset($_REQUEST['marker']) ? $_REQUEST['marker'] : '';
    $ply = isset($_REQUEST['ply']) ? $_REQUEST['ply'] : '';
    $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
   
  
  
  
    $_sql = "UPDATE OE_CUT_MU_DETAIL SET CUT_TABLE = '".$table."',
    WEIGHT_M2 = '".$weight_m2."',
    WEIGHT_YD = '".$weight_yd."',
    WIDTH_INCH = '".$width_inch."',
    LENGTH_YD = '".$length_yd ."',
    LENGTH_INCH = '".$length_inch ."',
    MARKER = '".$marker ."',
    PLY = '".$ply ."',
    UPDATE_BY = '".$username."',
    UPDATE_DATE = SYSDATE
    WHERE MU_SEQ = '".$mu_seq."' and GM_NO = '".$gm_no."'";
  
    $_sql2 = "COMMIT";
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql,$_sql2);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    } else {
  
      $result_out->status = (false);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });


  $app->post('/save_avg_end_code', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
    $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
    $avg_end_code = isset($_REQUEST['AVG_END_CODE']) ? $_REQUEST['AVG_END_CODE'] : '';
    $avg_end_cause = isset($_REQUEST['AVG_END_CAUSE']) ? $_REQUEST['AVG_END_CAUSE'] : '';
    $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
  
  
    $_sql = "UPDATE OE_CUT_MU_HEAD SET AVG_END_CODE = '".$avg_end_code."',
    AVG_END_CAUSE = '".$avg_end_cause."',
    UPDATE_BY = '".$username."',
    UPDATE_DATE = SYSDATE
    WHERE MU_SEQ = '".$mu_seq."' and SO_NO_DOC = '".$so_no_doc."'";
  
    $_sql2 = "COMMIT";
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql,$_sql2);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    } else {
  
      $result_out->status = (false);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });
  
  $app->post('/save_splice_loss_per', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
    $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
    $splice_loss_per_code = isset($_REQUEST['SPLICE_LOSS_CODE']) ? $_REQUEST['SPLICE_LOSS_CODE'] : '';
    $splice_loss_per_cause = isset($_REQUEST['SPLICE_LOSS_CAUSE']) ? $_REQUEST['SPLICE_LOSS_CAUSE'] : '';
    $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
  
  
    $_sql = "UPDATE OE_CUT_MU_HEAD SET SPLICE_LOSS_PER_CODE = '".$splice_loss_per_code."',
    SPLICE_LOSS_PER_CAUSE = '".$splice_loss_per_cause."',
    UPDATE_BY = '".$username."',
    UPDATE_DATE = SYSDATE
    WHERE MU_SEQ = '".$mu_seq."' and SO_NO_DOC = '".$so_no_doc."'";
  
    $_sql2 = "COMMIT";
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql,$_sql2);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    } else {
  
      $result_out->status = (false);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/save_total_cut_out_per', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
    $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
    $total_cut_out_per_code = isset($_REQUEST['TOTAL_CUT_OUT_PER_CODE']) ? $_REQUEST['TOTAL_CUT_OUT_PER_CODE'] : '';
    $total_cut_out_per_cause = isset($_REQUEST['TOTAL_CUT_OUT_PER_CAUSE']) ? $_REQUEST['TOTAL_CUT_OUT_PER_CAUSE'] : '';
    $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
  
  
    $_sql = "UPDATE OE_CUT_MU_HEAD SET TOTAL_CUT_OUT_PER_CODE = '".$total_cut_out_per_code."',
    TOTAL_CUT_OUT_PER_CAUSE = '".$total_cut_out_per_cause."',
    UPDATE_BY = '".$username."',
    UPDATE_DATE = SYSDATE
    WHERE MU_SEQ = '".$mu_seq."' and SO_NO_DOC = '".$so_no_doc."'";
  
    $_sql2 = "COMMIT";
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql,$_sql2);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    } else {
  
      $result_out->status = (false);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/save_remnants_loss', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
    $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
    $remnants_loss_code = isset($_REQUEST['REMNANTS_LOSS_CODE']) ? $_REQUEST['REMNANTS_LOSS_CODE'] : '';
    $remnants_loss_cause = isset($_REQUEST['REMNANTS_LOSS_CAUSE']) ? $_REQUEST['REMNANTS_LOSS_CAUSE'] : '';
    $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
  
  
    $_sql = "UPDATE OE_CUT_MU_HEAD SET REMNANTS_LOSS_CODE = '".$remnants_loss_code."',
    REMNANTS_LOSS_CAUSE = '".$remnants_loss_cause."',
    UPDATE_BY = '".$username."',
    UPDATE_DATE = SYSDATE
    WHERE MU_SEQ = '".$mu_seq."' and SO_NO_DOC = '".$so_no_doc."'";
  
    $_sql2 = "COMMIT";
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql,$_sql2);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    } else {
  
      $result_out->status = (false);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });

  $app->post('/save_widthloss', function (Request $request, Response $response) { // สร้าง route ขึ้นมารองรับการเข้าถึง url

    $mu_seq = isset($_REQUEST['MU_SEQ']) ? $_REQUEST['MU_SEQ'] : '';
    $so_no_doc = isset($_REQUEST['SO_NO_DOC']) ? $_REQUEST['SO_NO_DOC'] : '';
    $widthloss_code = isset($_REQUEST['WIDTHLOSS_CODE']) ? $_REQUEST['WIDTHLOSS_CODE'] : '';
    $widthloss_cause = isset($_REQUEST['WIDTHLOSS_CAUSE']) ? $_REQUEST['WIDTHLOSS_CAUSE'] : '';
    $username = isset($_REQUEST['username']) ? $_REQUEST['username'] : '';
  
  
    $_sql = "UPDATE OE_CUT_MU_HEAD SET WIDTH_LOSS_CODE = '".$widthloss_code."',
    WIDTH_LOSS_CAUSE = '".$widthloss_cause."',
    UPDATE_BY = '".$username."',
    UPDATE_DATE = SYSDATE
    WHERE MU_SEQ = '".$mu_seq."' and SO_NO_DOC = '".$so_no_doc."'";
  
    $_sql2 = "COMMIT";
    $result_out = new stdClass();
  
  
    $DataRows = ConnectDbnoAll($_sql,$_sql2);
  
    if ($DataRows !== false) {
  
  
      $result_out->status = (true);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    } else {
  
      $result_out->status = (false);
      $result_out->MU_SEQ = $mu_seq;
      $result_out->SQL = $_sql;
    }
  
    $SearchArray = [];
  
    $resultArray2 = array();
    array_push($resultArray2, $SearchArray);
    $response->getBody()->write(json_encode($result_out)); // สงคำตอบกลับ
  
    return $response; // ส่งคำตอบกลับร้า
  });
$app->run(); //สั่งระบบให้ทำงาน
