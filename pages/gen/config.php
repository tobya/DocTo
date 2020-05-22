<?php


 

  require "vendor/smarty/smarty/libs/Smarty.class.php";
  require "generation_functions.php";

$smarty =  new Smarty;

$smarty->template_dir = './templates';
$smarty->left_delimiter = '{[';
$smarty->right_delimiter = ']}';

