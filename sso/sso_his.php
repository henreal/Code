<?php
header("Expires:-1");
header("Cache-Control:no_cache");
header("Pragma:no-cache");

header("content-type:text/html;charset=utf-8");
$ssoYGDM = isset($_REQUEST["ygdm"])? $_REQUEST["ygdm"] : "";		//ygdm
$ssoPASS = isset($_REQUEST["pass"])? $_REQUEST["pass"] : "";		//pass
$ssoXTSB = isset($_REQUEST["xtsb"])? $_REQUEST["xtsb"] : "";		//xtsb

$ini_array = parse_ini_file("ssoconfig.ini");
$client = new SoapClient($ini_array["ssurl"]);		//接口地址
$ls_class = "SSO";						//请求类名，通过此接口检查员工是否登陆
$ls_action = "CheckUrl";				//请求方法
$ls_input = "{'ygdm':'". $ssoYGDM ."','pass':'" . $ssoPASS ."','xtsb':'" . $ssoXTSB . "'}";				//请求方法
$ls_output = "";						//返回JSON

$company = "hengrui";						//公司简称
$Token = "5465879";							//已授权Token
$TimeStamp = time();						//时间戳UTC/GMT 1970-01-01 00:00:00
$Message = $company . $Token . $TimeStamp . $ls_input;	//合并字串，规则：公司简称 + Token + 时间戳 + 入参
$Sign = md5($Message);						//MD5 对合并字串加密32位
$ls_code = $company . "_" . $TimeStamp . "|" . $Sign;		//生成ls_code，规则：公司简称 + "_" + 时间戳 + "|" + Sign

$param = array("ls_class"=>$ls_class, "ls_action"=>$ls_action, "ls_code"=>$ls_code, "ls_input"=>$ls_input, "ls_output"=>$ls_output);		//定义参数数组

$arr = $client->DoAction($param);			//使用接口的DoAction方法提交
$result = get_object_vars($arr);
if($result["DoActionResult"]==1){
	//echo $ssoYGDM . "工号";
    header("Refresh:1;url=/Desktop/Index.html?ssoLogin=true&chk=1&ygdm=" . $ssoYGDM);
}else{
    $ErrMsg = "SSO登陆失败！<br><a href=\"/Manage/Login.html\">帐号登陆</a>";
    echo $ErrMsg;
}
?>