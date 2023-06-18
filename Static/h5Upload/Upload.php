<?php
$img = isset($_POST['image'])? $_POST['image'] : '';
$imgDir = isset($_POST['FaceCard'])? $_POST['FaceCard'] : '';

// 获取图片
list($type, $data) = explode(',', $img);

// 判断类型
if(strstr($type,'image/jpeg')!==''){
	$ext = '.jpg';
}elseif(strstr($type,'image/gif')!==''){
	$ext = '.gif';
}elseif(strstr($type,'image/png')!==''){
	$ext = '.png';
}

// 生成的文件名
$imgFile = 'H'.time().$ext;
$imgPath = "/Upload" . "/" . $imgDir ;
$imgPath_S = $_SERVER['DOCUMENT_ROOT'] . str_replace("/","\\",$imgPath);
if (! is_dir ( $imgPath_S )) {mkdir( $imgPath_S, '0777',true );}    //建立目录
$photo = $imgPath_S . '\\' . $imgFile;
file_put_contents($photo, base64_decode($data), true);

// 返回
header('content-type:application/json;charset=utf-8');
$ret = array('url'=>$imgPath . "/" . $imgFile ,'info'=>'上传成功','status'=>'1');
echo json_encode($ret);
?>