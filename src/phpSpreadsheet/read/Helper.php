<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2019/5/21
 * Time: 15:24
 */

namespace phpSpreadsheetHelper\read;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use phpSpreadsheetHelper\SHelper;

class Helper extends SHelper
{
    /**
     * @var object Cached PhpSpreadsheet object
     */
    private $_objSpreadsheet;
    /**
     * @var SHelper
     */
    private $_objSheet;
    /**
     * @var maxColumn of sheet
     */
    private $columnCnt;
    /**
     * @var maxRow of sheet
     */
    private $_headRowCnt;
    /**
     * @var array tableData
     */
    private $_var=[
        'header'=>[],
        'data'=>[]
    ];
    /**
     * Helper constructor.
     */
    public function __construct(){
        $this->_objSpreadsheet = new Spreadsheet();
        $this->_objSheet = $this->_objSpreadsheet->getActiveSheet();
    }

    public function addFile($file_url,$upload_temp_path){
        $file_path = $this->downloadFile($file_url,$upload_temp_path);
        if (empty($file_path) OR !file_exists($file_path)) {
            throw new \Exception('文件不存在!');
        }
        $inputFileType = IOFactory::identify($file_path);
        $objRead = IOFactory::createReader($inputFileType);
        /* 建立excel对象 */
        $this->_objSpreadsheet = $objRead->load($file_path);
        @unlink($file_path);
//        empty($options) && $objRead->canRead($file_path);
        /* 获取指定的sheet表 */
        $this->_objSheet = $this->_objSpreadsheet->getSheet(0);


//        $data = $currSheet->toArray();
//        if (isset($options['mergeCells'])) {
//            /* 读取合并行列 */
//            $options['mergeCells'] = $this->_objSheet->getMergeCells();
//        }
        return $this;
    }

    /**
     * 设置最大列号
     * @param $columnCnt
     * @return $this
     */
    public function setColumn($columnCnt=0){
        if(is_numeric($columnCnt&&0 !== $columnCnt)){
            $this->columnCnt = $columnCnt;
        }else{
            $columnH = $this->_objSheet->getHighestColumn();
            $this->columnCnt = Coordinate::columnIndexFromString($columnH);
        }
        return $this;
    }

    public function getMaxColumn(){
        if(empty($this->columnCnt)){
            $columnH = $this->_objSheet->getHighestColumn();
            /* 兼容原逻辑，循环时使用的是小于等于 */
            $this->columnCnt = Coordinate::columnIndexFromString($columnH);
        }
        return $this->columnCnt;
    }

    public function setHeadRow($headRow){
        if(is_numeric($headRow)){
            $this->_headRowCnt = $headRow;
        }else{
            $this->_headRowCnt = 1;
        }
        return $this;
    }

    public function getHeadRow(){
        if(empty($this->_headRowCnt)){
            $this->_headRowCnt = 1;
        }
        return $this->_headRowCnt;
    }

    public function output(){
         $header = $this->_var['header'];
         if(empty($header)){
             return $data = $this->_objSheet->toArray();
         }

        $curr_header = $this->getHeadTitle();
        $dataRow = $this->_objSheet->getHighestDataRow();
        $headRow = $this->getHeadRow();
        $data = $this->_objSheet->toArray();
        $list = [];
        for($i=$headRow;$i<$dataRow;$i++){
            foreach ($header as $k=>$v){
                $key = array_search($v['name'],$curr_header);
                if($key!==false){
                    $list[$i-$headRow][$v['key']] = $this->getFormatValue($data[$i][$key],$v['format'],$v['name']);
                }
            }
        }
        return $list;
    }

    public function getFormatValue($value,$format,$remark=''){
        switch ($format){
            case  'FORMAT_TEXT':
                $value = trim($value);
                break;
            case 'FORMAT_NUMBER':
                if(!is_numeric($value)){
                    throw new Exception($remark.'必须为数字');
                }
                break;
        }
        return $value;
    }

    public function getHeadTitle(){
        $maxColumn = $this->getMaxColumn();
        $headRow = $this->getHeadRow();
        $data = $this->_objSheet->toArray();
        $curr_header = [];
        for($i=$headRow-1;$i<$headRow;$i++){
            foreach ($data[$i] as $v){
                $curr_header[]=$v;
            }
        }
        return $curr_header;
    }

    /**
     * 表头识别
     * @param $header
     * @return $this
     * @throws Exception
     */
    public function setHeader($header){
        if(empty($header)){
            throw new Exception('缺少表头识别');
        }
        foreach ($header as $v){
            $this->_var['header'][] = [
                'key'=>$v[0],
                'name'=>$v[1],
                'format'=>isset($v[2])?$v[2]:'FORMAT_TEXT'
            ];
        }
        return $this;
    }

    /**
     * @param $url
     * @param string $upload_temp_path 文件上传临时目录
     * @return string
     */
    private function downloadFile($url,$upload_temp_path)
    {
        $type = array_slice(explode('.', $url),-1,1);
        $file_name = date("YmdHis").mt_rand(1000, 9999).'.'.$type[0];
        //创建下载目录
        if (!is_dir($upload_temp_path)) {
            mkdir($upload_temp_path,0777, true);
        }
        $file_path = $upload_temp_path.'/'.$file_name;
//        if (!file_exists('./uploads/files')) {
//            mkdir('./uploads/files', 0777, true);
//        }
//        $file_path='./uploads/files/'.$file_name;
        $file = file_get_contents($url);
        file_put_contents($file_path,$file);
        return $file_path;
    }
}