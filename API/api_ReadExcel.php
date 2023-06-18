<?php
include "../Classes/PHPExcel/IOFactory.php";
$inputFileName = isset($_REQUEST['xlsFile'])? $_REQUEST['xlsFile'] : '';
$pType = isset($_REQUEST['type'])? $_REQUEST['type'] : '';
$inputFileName = $_SERVER['DOCUMENT_ROOT'] . $inputFileName;
$inputFileName = str_replace("/","\\", $inputFileName);
date_default_timezone_set("PRC");
// 读取excel文件
try {
    $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load($inputFileName);
} catch(Exception $e) {
    die("加载文件发生错误：".pathinfo($inputFileName,PATHINFO_BASENAME)." ".$e->getMessage());
}

// 读取sheet
$sheet = $objPHPExcel->getSheet(0);
$highestRow = $sheet->getHighestRow();
$highestColumn = $sheet->getHighestColumn();
//if($highestRow>100){$highestRow=20;}
// 获取一行的数据
$patten = array("\r\n", "\n", "\r");
$strData = "";
$strJson = "{\"code\":0,\"msg\":\"暂无数据（Excel）\",\"count\":". $highestRow .",\"data\":[";
$beginRow = 1;
if((int)$pType == 2){$beginRow=0;}
for ($row = $beginRow; $row <= $highestRow; $row++){
    // Read a row of data into an array
    $rowData = $sheet->rangeToArray("A" . $row . ":" . $highestColumn . $row, NULL, TRUE, FALSE);
    //这里得到的rowData都是一行的数据，得到数据后自行处理，我们这里只打出来看看效果
    if($row>$beginRow){ $strJson .= ","; $strData .= "@@";}
    $strJson .= "{\"fLen\":" . count($rowData[0]) . "";
        for($i=0;$i<count($rowData[0]);$i++){
			if($i>0){$strData .= "||";}
			$strJson .= ",\"VA".$i."\":\"" . str_replace($patten, "", $rowData[0][$i]) . "\"";
			$strData .= str_replace($patten, "", $rowData[0][$i]);
        }
		$strJson .= "}";
}
$strJson .= "]}";
if((int)$pType == 1 || (int)$pType == 2){
	echo $strData;
}else{
	echo $strJson;
}

?>