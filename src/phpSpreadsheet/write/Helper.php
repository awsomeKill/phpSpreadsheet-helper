<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2019/5/21
 * Time: 15:25
 */

namespace phpSpreadsheetHelper\write;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use phpSpreadsheetHelper\SHelper;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Helper extends SHelper
{
    /**
     * @var object Cached PhpSpreadsheet object
     */
    private $_objSpreadsheet;
    /**
     * @var object Cached Sheet object
     */
    private $_objSheet;
    /**
     * @var string Offset of Column
     */
    private $_columnOffset;
    /**
     * @var string max offset of Column
     */
    private $_columnMaxOffset;
    /**
     * @var string Offset of data
     */
    private $_dataOffset;
    /**
     * @var array Combined Data
     */
    private $_var = [
        'title'  =>  '',
        'header' =>  [] ,
        'data'  =>  [] ,
    ];


    public function __construct()
    {
        $this->_objSpreadsheet = new Spreadsheet();
        $this->_objSheet = $this->_objSpreadsheet->getActiveSheet();
        $this->_columnMaxOffset = $this->getMaximumColumn();
    }

    /**
     * @param $header
     * @return $this
     * @throws Exception
     */
    public function addHeader($header){
        if(empty($header)){
            throw new Exception('请填写表头信息');
        }
        for($i=0;$i<count($header);$i++){
            $v = $header[$i];
            $this->_var['header'][] = [
                'key' => $v[0],
                'width' => is_numeric($v[1])?$v[1]:'auto',
                'name' => $v[2],
                'format' => isset($v[3])?$v[3]:'default',  //格式
            ];
            if($i+1<count($header)){
                $this->_columnMaxOffset++;
            }
        }
        return $this;
    }

    private function setColumnWidth($value,$length){
        $cellWidth = $length == 'auto' ? strlen($value) : $length;
        $this->_objSheet->getColumnDimension($this->getMaximumColumn())->setWidth($cellWidth);
        return $this;
    }

    public function getMaximumColumn(){
        return $this->_objSheet->getHighestColumn();
    }

    public function getMaximumRow(){
        return $this->_objSheet->getHighestRow();
    }

    public function setData($expTableData){
        if(empty($expTableData)||!is_array($expTableData)){
            return $this;
        }
        foreach ($expTableData as $v){
            $this->_var['data'][] = $v;
        }
        return $this;
    }

    public function addTitle($title){
        $this->_var['title'] = $title;
        return $this;
    }


    public function output($filename,$extension='xlsx'){
        if(empty($extension)||!isset($this->_writeExtensions[strtoupper($extension)])){
            throw new Exception('缺少文件格式');
        }
        $this->combination();
        $_extension = $this->_writeExtensions[strtoupper($extension)]['extension'];
        $_contentType = $this->_writeExtensions[strtoupper($extension)]['contentType'];
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
        header("Content-Type:application/force-download");
        header("Content-Type:".$_contentType);
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");
        header('Content-Disposition:attachment;filename='.$filename.$_extension);
        header("Content-Transfer-Encoding:binary");
        header("Pragma: no-cache");
        $objWriter = new Xlsx($this->_objSpreadsheet);
        $objWriter->save('php://output');
        exit();
    }

    public function combination(){
        if(!empty($this->_var['title'])){
            $this->_objSheet->mergeCells('A1:'.$this->_columnMaxOffset.'1');
            $this->_objSheet->setCellValue('A1', $this->_var['title']);
            $this->_objSheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }

        $this->_columnOffset = $this->getMaximumColumn();
        $this->_dataOffset = $this->getMaximumRow();
        $this->_dataOffset++;
        for($i=0;$i<count($this->_var['header']);$i++){
            $this->_objSheet->setCellValue($this->_columnOffset.$this->_dataOffset, $this->_var['header'][$i]['name']);
            //设置单元格宽度
            $this->setColumnWidth($this->_var['header'][$i]['name'],$this->_var['header'][$i]['width']);
            //文本格式
            $this->setCellsType($this->_columnOffset,$this->_var['header'][$i]['format']);
            $this->_columnOffset++;
        }
        for($i=0;$i<count($this->_var['data']);$i++){
            $currentRowData = $this->_var['data'][$i];
            $__dataOffset= $this->_dataOffset+1+$i;
            $_currOffset = 0;
            for($col='A';$col != $this->_columnOffset;$col++){
                if (isset($currentRowData[$this->_var['header'][$_currOffset]['key']])) {
                    if (isset($this->_var['header'][$_currOffset]['format']) && $this->_var['header'][$_currOffset]['format'] == 'FORMAT_TEXT') {
                        $this->_objSheet->setCellValueExplicit($col.$__dataOffset, " ".$currentRowData[$this->_var['header'][$_currOffset]['key']], DataType::TYPE_STRING);
                    } elseif (isset($this->_var['header'][$_currOffset]['format']) && $this->_var['header'][$_currOffset]['format'] == 'FORMAT_NUMBER'){
                        $this->_objSheet->setCellValueExplicit($col.$__dataOffset, " ".$currentRowData[$this->_var['header'][$_currOffset]['key']], DataType::TYPE_NUMERIC);
                    }else {
                        $this->_objSheet->setCellValueExplicit($col.$__dataOffset, " ".$currentRowData[$this->_var['header'][$_currOffset]['key']], DataType::TYPE_STRING);
                    }
                }
                $_currOffset++;
            }
        }
        return $this;
    }

    public function setCellsType($cell,$format){
        switch ($format) {
            case 'FORMAT_NUMBER':
                $this->_objSheet->getStyle($cell)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER);
                break;
            case 'FORMAT_TEXT':
                $this->_objSheet->getStyle($cell)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_TEXT);
                break;
            default:
                $this->_objSheet->getStyle($cell)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_TEXT);
                break;
        }
    }
}