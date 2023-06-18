<?php
    header('content-type:text/html;charset=utf-8');
    error_reporting(0);
    require "./incCaptcha.class.php";  //先把类包含进来，实际路径根据实际情况进行修改。
    $Action = isset($_REQUEST["Action"]) ? $_REQUEST["Action"] : '';
    $tcode = isset($_REQUEST["code"]) ? $_REQUEST["code"] : '';
    $errmsg = "{\"err\":true,\"errcode\":500,\"errmsg\":\"no\"}";
    session_start();
    if($Action=="GetCode"){
        $tcode = isset($_REQUEST["code"]) ? $_REQUEST["code"] : '';
        $tmpCaptcha = $_SESSION["authnum"];
        if(strtolower($tcode)==$tmpCaptcha){            //转为小写字母验证
            if($tcode<>""){$errmsg="{\"err\":false, \"errcode\":0,\"errmsg\":\"ok\"}";}
        }else{
            $errmsg="{\"err\":true, \"errcode\":500,\"errmsg\":\"" . $tcode . "\"}";
        }
        echo $errmsg;
    }else{
        ob_clean();                     //清除缓存
        $_vc = new ValidateCode();      //实例化一个对象
        $_vc->doimg();  
        $_SESSION["authnum"] = $_vc->getCode(); //验证码保存到SESSION中
        //以下是将缓存写入文本中，用于跨域访问
        $session_text = '../Upload/SESSION/' . session_id() . '.txt';
        $file = fopen($session_text, 'w');
        $str = $_SESSION['authnum'];//内容
        $res = fwrite($file, $str);	//写入
        fclose($file);	//关闭fo
    }
?>