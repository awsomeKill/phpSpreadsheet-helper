<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2019/5/20
 * Time: 11:49
 */

require_once  __DIR__.'/vendor/autoload.php';

use phpSpreadsheetHelper\SHelper;

$file_path = __DIR__.'/list1.xls';

$list =  SHelper::make('read')
    ->addFile($file_path)
    ->setHeader([
        ['product_name','商品名称'],
        ['unit_name','单位'],
        ['category_name','商品分类'],
        ['price','价格'],
        ['cost_price','成本价'],
        ['bar_code','商品条码'],
        ['stock','库存']
    ])
    ->output();
  var_dump($list);die;
