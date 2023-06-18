<?php
header("Expires:-1");
header("Cache-Control:no_cache");
header("Pragma:no-cache");

header("content-type:text/html;charset=utf-8");
$getType = isset($_REQUEST["Type"])? $_REQUEST["Type"] : "";
$ssoYGDM = isset($_REQUEST["YGDM"])? $_REQUEST["YGDM"] : "";		//ygdm
$ssoPASS = isset($_REQUEST["PASS"])? $_REQUEST["PASS"] : "";		//pass
$ssoXTSB = isset($_REQUEST["XTSB"])? $_REQUEST["XTSB"] : "";		//xtsb
$ssoMsg = isset($_REQUEST["MSG"])? $_REQUEST["MSG"] : "";			//发送消息内容
$isReply = isset($_REQUEST["isReply"])? $_REQUEST["isReply"] : "";		//消息是否回复

$ini_array = parse_ini_file("../sso/ssoconfig.ini");	//取接口配置
$apiURL = $ini_array["ssurl"];		//接口地址

$ls_input = "";		//请求参数(JSON)
$client = new SoapClient($apiURL);			//接口地址
if($getType == "CheckUrl"){
	$ls_class = "SSO";						//请求类名，通过此接口检查员工是否登陆
	$ls_action = "CheckUrl";				//请求方法
	$ls_input = "{'ygdm':'". $ssoYGDM ."','pass':'". $ssoPASS ."','xtsb':'". $ssoXTSB ."'}";				//请求方法
}else if($getType == "GetAllRyxx"){
	$ls_class = "RS_RSDA";
	$ls_action = $getType;
}else if($getType == "GetRyxxForJXGL"){
	$ls_class = "RS_YGDA";
	$ls_action = $getType;
}else if($getType == "SendMSG"){			//发送消息
	$ls_class = "GY_SMS";
	$ls_action = "SendQYWXByYgdmImmOut";
	$ls_input = "{'ygdm':'". $ssoYGDM ."','content':'". $ssoMsg ."','isreply':'". $isReply ."'}";				//请求方法
}else{
	$ls_class = "Basic";
	$ls_action = $getType;
}

$ls_output = "";							//返回JSON
$company = "hengrui";						//公司简称
$Token = "5465879";							//已授权Token
$TimeStamp = time();						//时间戳UTC/GMT 1970-01-01 00:00:00
$Message = $company . $Token . $TimeStamp . $ls_input;	//合并字串，规则：公司简称 + Token + 时间戳

$Sign = md5($Message);						//MD5 对合并字串加密32位
$ls_code = $company . "_" . $TimeStamp . "|" . $Sign;		//生成ls_code，规则：公司简称 + "_" + 时间戳 + "|" + Sign

$param = array("ls_class"=>$ls_class, "ls_action"=>$ls_action, "ls_code"=>$ls_code, "ls_input"=>$ls_input);		//定义参数数组

$arr = $client->DoAction($param);			//使用接口的DoAction方法提交
$result = get_object_vars($arr);

$errNum = $result["DoActionResult"];
$strContents = "";
$nowTime = date("Y-m-d H:i:s");
$putFile = "";
if($getType=="GetKsDict"){
	if($errNum == "0"){
		$strContents = "{\"Return\": false,\"Err\": 500,\"reMessge\": \"连接远程科室接口数据失败！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":[]}";
	}else{
		$strContents = "{\"Return\": true,\"Err\": 0,\"reMessge\": \"科室接口数据连接成功！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":". $result["ls_output"]."}";
	}
	$putFile = "../Upload/Department.txt";
	file_put_contents($putFile, $strContents);
	echo $strContents;
}else if($getType=="GetAllRyxx"){
	if($errNum == "0"){
		$strContents = "{\"Return\": false,\"Err\": 500,\"reMessge\": \"连接远程全部人员信息失败！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":[]}";
		echo $strContents;
	}else{
		$strContents = "{\"Return\": true,\"Err\": 0,\"reMessge\": \"连接远程全部人员成功！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":". $result["ls_output"]."}";
		echo "{\"Return\": true,\"Err\": 0,\"reMessge\": \"远程全部人员连接成功！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":[]}";
	}
	$putFile = "../Upload/AllTeacher.txt";
	file_put_contents($putFile, $strContents);
}else if($getType=="GetYgDict"){
	if($errNum == "0"){
		$strContents = "{\"Return\": false,\"Err\": 500,\"reMessge\": \"连接远程员工接口数据失败！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":[]}";
		echo $strContents;
	}else{
		$strContents = "{\"Return\": true,\"Err\": 0,\"reMessge\": \"员工接口数据连接成功！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":". $result["ls_output"]."}";
		echo "{\"Return\": true,\"Err\": 0,\"reMessge\": \"员工接口数据连接成功！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":[]}";
	}
	$putFile = "../Upload/Teacher.txt";
	file_put_contents($putFile, $strContents);
}else if($getType=="GetRyxxForJXGL"){
	if($errNum == "0"){
		$strContents = "{\"Return\": false,\"Err\": 500,\"reMessge\": \"获取远程员工数据失败！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":[]}";
		echo $strContents;
	}else{
		$strContents = "{\"Return\": true,\"Err\": 0,\"reMessge\": \"获取远程员工数据成功！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":". $result["ls_output"]."}";
		echo "{\"Return\": true,\"Err\": 0,\"reMessge\": \"员工接口数据连接成功！\",\"ReStr\": \"Update:" . $nowTime . "\",\"reData\":[]}";
	}
	$putFile = "../Upload/Teacher08.txt";
	file_put_contents($putFile, $strContents);
}else if($getType=="CheckUrl"){
	session_start();
	if($errNum == 1){
		$_SESSION['CheckUrl'] = "OK";
		echo "OK<br>";
		//header("Refresh:3;url=/API/SSO.html?ssoLogin=true&chk=1&ygdm=" . $ssoYGDM);
	}else{
		echo "NO<br>";
		$_SESSION['CheckUrl'] = "NO";
		//header("Refresh:3;url=/API/SSO.html?ssoLogin=true&chk=0&ygdm=" . $ssoYGDM);
	}
	$strContents = "{\"code\": 100,\"Err\":\"" . $result["DoActionResult"] . "\",\"reMessge\":\"员工接口认证通讯成功！\",\"Update\":\"" . $nowTime . "\",\"ygdm\":\"" . $ssoYGDM . "\",\"pass\":\"" . $ssoPASS . "\",\"reData\":\"\"}";
	echo $strContents;
	$putFile = "../Upload/ChkTeacher.txt";
	file_put_contents($putFile, $strContents);
}else if($getType=="SendMSG"){
	$strContents = "{\"code\": 100,\"Err\":\"1\",\"reMessge\":\"发送消息失败！\",\"Update\":\"" . $nowTime . "\",\"ygdm\":\"" . $ssoYGDM . "\",\"reData\":". $result["ls_output"]."}";
	echo $arr;
	$putFile = "../Upload/SendMSG.txt";
	file_put_contents($putFile, $strContents);
}
?>