<?php
require_once  __DIR__.'/../vendor/autoload.php';

use phpSpreadsheetHelper\SHelper;


$file_path = __DIR__.'/../list1.csv';

$upload_temp_path = __DIR__.'/public/uploads/temp';
$list =  SHelper::make('read')
    ->addFile($file_path,$upload_temp_path)
    ->setHeader([
        ['name','商品名称'],
        ['num','数量'],
        ['category','分类'],
        ['price','售价'],
    ])
    ->output();

