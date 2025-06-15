<?php


    if (!function_exists('docto_path')){
    function docto_path($path){
       return base_path('../' . $path);
    }
}
