<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2019/5/17
 * Time: 10:54
 */

namespace phpSpreadsheetHelper;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class SHelper{


    /**
     * @var object Cached PhpSpreadsheet Sheet object
     */
    private static $_objSheet;

    /**
     * @var  int Current row offset for the actived sheet
     */
    private static $_offsetRow;

    /**
     * @var int Current column offset for the actived sheet
     */
    private static $_offsetCol;

    /**
     * @var array Extensions for read
     */
    private static $_readExtensions = [
        'OASIS' => '.ods',
        'Excel2007' => '.xlsx',
        'Excel97' => '.xls',
        'Excel95' => '.xls',
        'Excel2003' => '.xml',
        'HTML' => '.html',
        'SYLK' => '.sylk',
        'Gnumeric' => '.gnumeric',
        'CSV' => '.csv'
    ];

    /**
     * @var array Extensions for write
     */
    protected $_writeExtensions = [
        'OASIS' => [
            'extension' => '.ods',
            'contentType' => 'application/vnd.oasis.opendocument.spreadsheet'
        ],
        'XLSX' =>[
            'extension'=>'.xlsx',
            'contentType'=> 'application/vnd.ms-excel'
        ],
        'XLS' => [
            'extension' => '.xls',
            'contentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ],
        'HTML' => [
            'extension' => '.html',
            'contentType' => 'text/html'
        ],
        'CSV' => [
            'extension' => '.csv',
            'contentType' => 'text/csv'
        ],
        'PDF' => [
            'extension' => '.pdf',
            'contentType' => 'text/pdf'
        ]
    ];

    /**
     * 初始化
     */
    public function initialize()
    {}


    public static function make($action = '')
    {
        if($action == ''){
            throw new Exception('没有指定操作');
        }else {
            $action = strtolower($action);
        }

        // 构造器类路径
        $class = 'phpSpreadsheetHelper\\'. $action .'\\Helper';
        if (!class_exists($class)) {
            throw new Exception($action . '构建器不存在', 8002);
        }
        return new $class;
    }





}