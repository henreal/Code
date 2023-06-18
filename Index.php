<?php
header("Expires:-1");
header("Cache-Control:no_cache");
header("Pragma:no-cache");
//echo "调试中";exit();
$strChk = isMobile()?"移动端":"PC平台";


/**以下为跳转**/
if(isMobile()){
    header("Location:/Touch/Index.html");
}else{
    header("Location:/Desktop/Index.html");
}

function isMobile() {
    if (isset($_SERVER['HTTP_X_WAP_PROFILE'])) {return true;}
    if (isset($_SERVER['HTTP_VIA'])) { return stristr($_SERVER['HTTP_VIA'], "wap") ? true : false;}
    if (isset($_SERVER['HTTP_USER_AGENT'])) {
        $clientkeywords = array('nokia','sony','ericsson','mot','samsung','htc','sgh','lg','sharp','sie-','philips','panasonic','alcatel','lenovo','iphone','ipod','blackberry','meizu','android','netfront','symbian','ucweb','windowsce','palm','operamini','operamobi','openwave','nexusone','cldc','midp','wap','mobile','MicroMessenger');
        if (preg_match("/(" . implode('|', $clientkeywords) . ")/i", strtolower($_SERVER['HTTP_USER_AGENT']))) { return true;}
    }
    if (isset ($_SERVER['HTTP_ACCEPT'])) {
        if ((strpos($_SERVER['HTTP_ACCEPT'], 'vnd.wap.wml') !== false) && (strpos($_SERVER['HTTP_ACCEPT'], 'text/html') === false || (strpos($_SERVER['HTTP_ACCEPT'], 'vnd.wap.wml') < strpos($_SERVER['HTTP_ACCEPT'], 'text/html')))) { return true;}
    }
    return false;
}
?>