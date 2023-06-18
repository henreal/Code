<?php
/**
 * Layui upload
 */

//取得上传文件信息
$fileName = $_FILES["file"]["name"];
$fileType = $_FILES["file"]["type"];
$fileError = $_FILES["file"]["error"];
$fileSize = $_FILES["file"]["size"];
$tempName = $_FILES["file"]["tmp_name"];      //临时文件名
$uploadDir = isset($_REQUEST["UploadDir"]) ? $_REQUEST["UploadDir"] : '';
$uploadDir = !empty($uploadDir) ? "/".$uploadDir : $uploadDir;
$uploadDir = "/Upload" . $uploadDir;
$uploadDir_s = $_SERVER['DOCUMENT_ROOT'] . str_replace("/","\\",$uploadDir);
if (! is_dir ( $uploadDir_s )) {mkdir( $uploadDir_s, '0777',true );}    //建立目录

$fileExtName = substr($fileName,strrpos($fileName,".") + 1);    //上传后文件名
$newFileName = "HR-". time() . randomNum(6). "." . $fileExtName;
$ret = array("code"=>"500","msg"=>"","data"=>array("oldFileName"=>$fileName,"src"=>"","fileType"=>$fileType,"fileSize"=>$fileSize,"maxSize"=>"30M","fileExtName"=>$fileExtName,"UploadDir"=>$uploadDir));
//定义上传文件类型
$typeList = array("image/jpeg","image/jpg","image/png","image/gif","audio/mpeg","application/vnd.ms-excel","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","application/msword","application/vnd.openxmlformats-officedocument.wordprocessingml.document","text/plain","application/x-zip-compressed","application/pdf");    //定义允许的类型

if($fileError>0){
    //上传文件错误编号判断
    switch ($fileError) {
        case 1:
            $message="上传的文件超过了php.ini 中 upload_max_filesize 选项限制的值。";
            break;
        case 2:
            $message="上传文件的大小超过了 HTML 表单中 MAX_FILE_SIZE 选项指定的值。";
            break;
        case 3:
            $message="文件只有部分被上传。";
            break;
        case 4:
            $message="没有文件被上传。";
            break;
        case 6:
            $message="找不到临时文件夹。";
            break;
        case 7:
            $message="文件写入失败";
            break;
        case 8:
            $message="由于PHP的扩展程序中断了文件上传";
            break;
    }
    $ret["msg"] = $message;
    exit(json_encode($ret));
}
if(!is_uploaded_file($tempName)){
    //判断是否是POST上传过来的文件
    $ret["msg"] = "不是通过HTTP POST方式上传上来的";
    exit(json_encode($ret));
}else{
    if(!in_array($fileType, $typeList)){
        $ret["msg"] = "上传的文件不是指定类型：". $fileType;
        exit(json_encode($ret));
    }else{
/**        if(!getimagesize($tempName)){
            //避免用户上传恶意文件,如把病毒文件扩展名改为图片格式
            $ret["msg"] = "上传的文件不是图片";
            exit(json_encode($ret));
        }*/
    }
    if($fileSize>30000000){
        //对特定表单的上传文件限制大小
        $ret["msg"] ="上传文件超出限制大小";
        exit(json_encode($ret));
    }else{
        //避免上传文件的中文名乱码
        $fileName=iconv("UTF-8", "GBK", $fileName);//把iconv抓取到的字符编码从utf-8转为gbk输出
        $fileName=str_replace(".", time().".", $fileName);//在图片名称后加入时间戳，避免重名文件覆盖
        if(move_uploaded_file($tempName, $uploadDir_s ."/".$newFileName)){
            $ret["code"] ="0";$ret["msg"] = "上传文件成功";$ret["data"]["src"] =  $uploadDir . "/".$newFileName;
        }else{
            $ret["msg"] = "上传文件失败";
        }
        exit(json_encode($ret));
    }
}

//生成随机文件名函数
function randomNum($length){
    $hash = '';
    $chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz';
    $max = strlen($chars) - 1;
    mt_srand((double)microtime() * 1000000);
    for($i = 0; $i < $length; $i++){
        $hash .= $chars[mt_rand(0, $max)];
    }
    return $hash;
}
?>