<?php

/* Connect Database */
define("BASE_HATCH_HOST", "(DESCRIPTION=(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.79)(PORT = 1521)))(CONNECT_DATA=(SID=NYTG)))");
define("BASE_HATCH_USER", "HATCH");
define("BASE_HATCH_PASSWORD", "HATCH");

define("BASE_RTS_HOST", "(DESCRIPTION=(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.6.79)(PORT = 1521)))(CONNECT_DATA=(SID=NYTG)))");
define("BASE_RTS_USER", "RTS");
define("BASE_RTS_PASSWORD", "RTS");


/* Set Lang */
date_default_timezone_set("Asia/Bangkok");



function ConnectedDB() //ส่วนของการเลือก baselocation
{
    return oci_connect(BASE_HATCH_USER, BASE_HATCH_PASSWORD, BASE_HATCH_HOST, 'AL32UTF8');
}





?>

