<?php

require "config.php";

require "loadcommands.php";


$smarty->assign('Params',$Explain);

foreach ($Commands as $CommandName => $Command) {


foreach ($Command['Items'] as $keytag => $Item) {
    # code...

$smarty->assign('Command', $Item);
    # code...

$MDFile = $smarty->fetch($Command['Template']);

if (!file_exists('../all/')){

mkdir('../all/');
}
$Fn =   $CommandName . @$Item['FileTypeTitleExtra'] . $Item['FileTypeExt'] . '.md' ;
$Links[] = $Fn;
echo "Create File : $Fn\n";
file_put_contents('../all/' . $Fn, $MDFile);

}
}

$md = '
List of Files
==

';
foreach ($Links as $key => $value) {
    $md .= " - [$value]($value)\n";
}
file_put_contents('../all/index.md' , $md);
echo "\nList index.md";

