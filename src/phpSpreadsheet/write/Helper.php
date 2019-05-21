<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2019/5/21
 * Time: 15:25
 */

namespace phpSpreadsheetHelper\write;


use phpSpreadsheetHelper\SHelper;

class Helper extends SHelper
{
    public function __construct()
    {
        self::$_objSpreadsheet = new Spreadsheet();
    }
}