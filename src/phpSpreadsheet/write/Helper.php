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
     * @var  object Cached Sheet object
     */
    private $_objSheet;
    /**
     * @var array Cached Sheet header
     */
    private $expCellName;
    /**
     * @var integer Maximum column number of header
     */
    private $cellNum;

    /**
     * @var array title of
     */
    private  $cellName = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ','DA','DB','DC','DD','DE','DF','DG','DH','DI','DJ','DK','DL','DM','DN','DO','DP','DQ','DR','DS','DT','DU','DV','DW','DX','DY','DZ','EA','EB','EC','ED','EE','EF','EG','EH','EI','EJ','EK','EL','EM','EN','EO','EP','EQ','ER','ES','ET','EU','EV','EW','EX','EY','EZ','FA','FB','FC','FD','FE','FF','FG','FH','FI','FJ','FK','FL','FM','FN','FO','FP','FQ','FR','FS','FT','FU','FV','FW','FX','FY','FZ','GA','GB','GC','GD','GE','GF','GG','GH','GI','GJ','GK','GL','GM','GN','GO','GP','GQ','GR','GS','GT','GU','GV','GW','GX','GY','GZ','HA','HB','HC','HD','HE','HF','HG','HH','HI','HJ','HK','HL','HM','HN','HO','HP','HQ','HR','HS','HT','HU','HV','HW','HX','HY','HZ'];


    public function __construct()
    {
        $this->_objSpreadsheet = new Spreadsheet();
        $this->_objSheet = $this->_objSpreadsheet->getActiveSheet();
    }


    public function addHeader($header){
        $this->cellNum  = count($header);
        $this->expCellName = $header;
        for($i=0;$i<$this->cellNum;$i++){
            $this->_objSheet->setCellValue($this->cellName[$i].'1', $this->expCellName[$i][2]);
            //根据内容设置单元格宽度
            $cellWidth = $this->expCellName[$i][1] == 'auto' ? strlen($this->expCellName[$i][2]) : $this->expCellName[$i][1];
            $this->_objSheet->getColumnDimension($this->cellName[$i])->setWidth($cellWidth);

            $this->_objSheet->getStyle($this->cellName[$i])->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_TEXT);
            if(isset($expCellName[$i][3])) {
                switch ($expCellName[$i][3]) {
                    case 'FORMAT_NUMBER':
                        $this->_objSheet->getStyle($this->cellName[$i])->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER);
                        break;
                    case 'FORMAT_TEXT':
                        $this->_objSheet->getStyle($this->cellName[$i])->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_TEXT);
                        break;
                }
            }
        }
        return $this;
    }

    public function setData($expTableData){
        $dataNum  = count($expTableData);
        for($i=0;$i<$dataNum;$i++){
            for($j=0;$j<$this->cellNum;$j++){
                if (isset($expTableData[$i][$this->expCellName[$j][0]])) {
                    if (isset($this->expCellName[$j][3]) && $this->expCellName[$j][3] == 'FORMAT_TEXT') {
                        $this->_objSheet->setCellValueExplicit($this->cellName[$j].($i+2), " ".$expTableData[$i][$this->expCellName[$j][0]], DataType::TYPE_STRING);
                    } elseif (isset($this->expCellName[$j][3]) && $this->expCellName[$j][3] == 'FORMAT_NUMBER'){
                        $this->_objSheet->setCellValueExplicit($this->cellName[$j].($i+2), " ".$expTableData[$i][$this->expCellName[$j][0]], DataType::TYPE_NUMERIC);
                    }else {
                        $this->_objSheet->setCellValueExplicit($this->cellName[$j].($i+2), " ".$expTableData[$i][$this->expCellName[$j][0]], DataType::TYPE_STRING);
                    }
                }
            }
        }
        return $this;
    }


    public function output($filename,$extension='xlsx'){
        if(empty($extension)||!isset($this->_writeExtensions[strtoupper($extension)])){
            throw new Exception('缺少文件格式');
        }
        $extension = $this->_writeExtensions[strtoupper($extension)]['extension'];
        $contentType = $this->_writeExtensions[strtoupper($extension)]['contentType'];
        //下载
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
        header("Content-Type:application/force-download");
        header("Content-Type:".$contentType);
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");
        header('Content-Disposition:attachment;filename='.$filename.$extension);
        header("Content-Transfer-Encoding:binary");
        header("Pragma: no-cache");
        $objWriter = new Xlsx($this->_objSpreadsheet);
        $objWriter->save('php://output');
        exit();
    }
}