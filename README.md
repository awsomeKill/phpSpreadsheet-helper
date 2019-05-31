# phpSpreadsheet-helper
Created for easier import and export

Run Composer in your project:

      composer require awsome-kill/php-spreadsheet-helper



### Export Key-value pair array


```php
return SHelper::make('write')
    ->addTitle('XX商品列表')
    ->addHeader([
        ['name',20,'商品名称'],
        ['num',15,'数量'],
        ['category',20,'分类'],
        ['price',15,'售价']
    ])
    ->setData($data)
    ->output('商品表','xlsx');
```
<img src="https://github.com/awsomeKill/phpSpreadsheet-helper/blob/master/img/export.png"/>


### Import into a specified key-value pair array

```php
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
```




