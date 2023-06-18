<?php
session_start();
$session_text = session_id() . '.txt';
var_dump($session_text);
echo "<br>";
echo "" . $_SESSION['authnum'];
$file = fopen($session_text, 'w'); //打开文件
$str = $_SESSION['authnum'];//内容
$res = fwrite($file, $str);	//写入
fclose($file);	//关闭fo

/* //session_start();	//必须加，这样才可以使用session
$reTips = ['err'=>true, 'errcode'=>500, 'errmsg'=>'验证失败！_dd', 'icon'=>2, 'id'=>0,];
$s = isset($_SESSION['authnum']);	//判断SESSION是否存在
$captcha = json_encode($_POST);
var_dump($s); */
/* if($s){
	echo $_SESSION['authnum'];
}else{
	echo json_encode($reTips);
} */

//PDO连接mysql
/* try{
	$pdo = new PDO("mysql:host=127.0.0.1;port=3306;dbname=hrcms","root", "123456");	//初始化PDO
	echo '连接成功';
}catch (PDOException $e){
	echo 'Error:' . $e->getMessage() . '<br>';
} */

//mysqli连接
/* try{
	$conn = mysqli_connect("127.0.0.1:3307", "root", "123456", "hrcms");	//建立mysql连接
}catch($e){
	echo '失败';
}

if(!$conn){
	echo '连接失败：' . mysqli_connect_error;
} */

?>
