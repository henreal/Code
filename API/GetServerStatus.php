<?php
header("Access-Control-Allow-Origin:*");
header("Access-Control-Allow-Methods:POST");
header("Access-Control-Allow-Headers:x-requested-with,content-type");

// 本文件用于检查服务器是否在线，请勿删除！
// 返回格式必须为JSON，参数：reStatus/状态，布尔值　reMSG/提示消息

$date = date_create(date("Y-m-d H:i:s"));
date_add($date, date_interval_create_from_date_string("24 hours"));
$reJsonStr = "{\"reStatus\": true, \"reMSG\":\"系统运行正常!\"}";
echo $reJsonStr;
?>
