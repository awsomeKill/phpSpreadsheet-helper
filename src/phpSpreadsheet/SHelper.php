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
     * @var object Cached PhpSpreadsheet object
     */
    private static $_objSpreadsheet;

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
    private static $_writeExtensions = [
        'OASIS' => '.ods',
        'Excel2007' => '.xlsx',
        'Excel97' => '.xls',
        'HTML' => '.html',
        'CSV' => '.csv',
        'PDF' => '.pdf'
    ];


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

    public static function resetSpread(){
        self::$_objSheet = NULL;
        self::$_offsetRow = 0;
        self::$_offsetCol = 0; // A1 => 1
    }

}