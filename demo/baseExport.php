<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2019/5/22
 * Time: 17:31
 */

require_once  __DIR__.'/../vendor/autoload.php';

use phpSpreadsheetHelper\SHelper;

$data=[
    ['name'=>'羊肉串','num'=>'14','category'=>'烧烤','price'=>'3/10'],
    ['name'=>'蔬菜沙拉','num'=>'28','category'=>'凉菜','price'=>'15'],
    ['name'=>'牛肉串','num'=>'50','category'=>'烧烤','price'=>'9'],
    ['name'=>'酸奶','num'=>'20','price'=>'3.5','category'=>'奶制品']
];

return SHelper::make('write')
    ->addTitle('XX商品列表')
    ->addHeader([
        ['name',20,'商品名称'],
        ['num',15,'数量'],
        ['category',20,'分类'],
        ['price',15,'售价']
    ])
    ->setData($data)
    ->output('商品表','csv');