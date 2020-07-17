<?php
header("Access-Control-Allow-Origin: *");
header("content-type:text/html;charset=utf-8");
$request_type = $_SERVER['REQUEST_METHOD'];

if($request_type == 'GET'){
    header("Location: index.html");
    exit();
}

/*读取excel文件，并进行相应处理*/
$fileName = "财务计划.xlsx";
if (!file_exists($fileName)) {
    exit("文件".$fileName."不存在");
}
$startTime = time(); //返回当前时间的Unix 时间戳
require_once './classes/PHPExcel/IOFactory.php';//引入PHPExcel类库
$objPHPExcel = PHPExcel_IOFactory::load($fileName);//获取sheet表格数目
$sheetCount = $objPHPExcel->getSheetCount();//默认选中sheet0表
$sheetSelected = 0;$objPHPExcel->setActiveSheetIndex($sheetSelected);//获取表格行数
$rowCount = $objPHPExcel->getActiveSheet()->getHighestRow();//获取表格列数
$columnCount = $objPHPExcel->getActiveSheet()->getHighestColumn();
// echo "<div>Sheet Count : ".$sheetCount."　　行数： ".$rowCount."　　列数：".$columnCount."</div>";
$dataArr = array();
$tmp_obj = NULL;
/* 循环读取每个单元格的数据 */
//行数循环
for ($row = 2; $row <= $rowCount; $row++){
    //列数循环 , 列数是以A列开始
    for ($column = 'A'; $column <= $columnCount; $column++) {
        $tmp_obj[$column] = $objPHPExcel->getActiveSheet()->getCell($column.$row)->getValue();
    }
    $dataArr[] = $tmp_obj;
    $tmp_obj = NULL;
    // echo "<br/>消耗的内存为：".(memory_get_peak_usage(true) / 1024 / 1024)."M";
    // $endTime = time();
    // echo "<div>解析完后，当前的时间为：".date("Y-m-d H:i:s")."总共消耗的时间为：".(($endTime - $startTime))."秒</div>";
    //var_dump($dataArr);
    //$dataArr = NULL;
}
$_res['code'] = 1;
$_res['data'] = $dataArr;
print_r(json_encode($_res));
?>