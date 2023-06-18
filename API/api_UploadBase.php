<?php
/**
 * upload Base64 Data
 */
$strBase64 = isset($_REQUEST["formFile"]) ? $_REQUEST["formFile"] : '';
$uploadDir = isset($_REQUEST["UploadDir"]) ? $_REQUEST["UploadDir"] : '';
$uploadDir = !empty($uploadDir) ? "/".$uploadDir : $uploadDir;
$uploadDir = "/Upload" . $uploadDir;
$uploadDir_s = $_SERVER['DOCUMENT_ROOT'] . str_replace("/","\\",$uploadDir);
if (! is_dir ( $uploadDir_s )) {mkdir( $uploadDir_s, '0777',true );}    //建立目录
//echo $strBase64;

echo base64_image_content($strBase64, $uploadDir_s);

/**
 * 将Base64图片转换为本地图片并保存
 * @param  [Base64] $base64_image_content [要保存的Base64]
 * @param  [目录] $path [要保存的路径]
 */
function base64_image_content($base64_image_content,$path){
	global $uploadDir;
    //匹配出图片的格式
    if (preg_match('/^(data:\s*image\/(\w+);base64,)/', $base64_image_content, $result)){
        $type = str_replace("jpeg","jpg",$result[2]);
        $new_path = "/".date('Ymd',time())."/";
        if(!file_exists($path.$new_path)){		//检查是否有该文件夹，如果没有就创建，并给予最高权限
            mkdir($path.$new_path, 0700);
		}
		$new_fileName = "HR" . time() . ".{$type}";
		$new_file = $path.$new_path.$new_fileName;
		
        if (file_put_contents($new_file, base64_decode(str_replace($result[1], '', $base64_image_content)))){
            return $uploadDir . $new_path . $new_fileName;
        }else{
            return false;
        }
    }else{
        return false;
    }
}

?>