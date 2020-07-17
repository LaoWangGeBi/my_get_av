<?php
header("Access-Control-Allow-Origin: *");
header("content-type:text/html;charset=utf-8");
$request_type = $_SERVER['REQUEST_METHOD'];

if($_SERVER['REQUEST_URI'] == '/'){
    header("Location: index.html");
    exit();
}
if($request_type == 'GET'){
    $_receive = $_GET;
}else{
    $_receive = $_POST;
}
$key = trim($_receive['key']);
$key = empty($key) ? '' : $key;
$page = trim($_receive['page']);
$page = empty($page) ? 1 : $page;
$page_size = 18;

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
$_data_len = count($dataArr);
$_count = floor($_data_len/$page_size);
$_count = $_data_len%$page_size > 0 ? $_count+1 : $_count;
if($page == $_count ){
    if($page == 1){
        $_new_list = $dataArr;
    }else{
        $_new_list = array_slice($dataArr,($page-1)*$page_size-1);
    }
}else{
    if($page == 1){
        $_new_list = array_slice($dataArr,0,$page_size);
    }else{
        $_new_list = array_slice($dataArr,($page-1)*$page_size-1,$page_size);
    }
}
$_res['code'] = 1;
$_res['current'] = $page;
$_res['count'] = $_count;
if($key){
    $_res['key'] = $key;
}
$_res['data'] = $_new_list;
print_r(json_encode($_res));
?>
