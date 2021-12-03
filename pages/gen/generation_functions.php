<?php

function LoadResourceFile($fn) {
    $fullfn = __DIR__ . '\\..\\..\\src\\res\\' . $fn;

    return file_get_contents($fullfn); 

}

function LoadResourceFiles(){
    $allfiles = scandir(__DIR__ . '\\..\\..\\src\\res\\');

    foreach ($allfiles as $key => $fn) {
        
        if (strlen($fn) > 2 ){ // ignore . ..
                //echo $fileinfo['filename'] . ':' . strlen($fileinfo['filename']);
            if (strpos( $fn , '__history') !== false){continue;}
            $info = pathinfo($fn);

            $AllContent[$info['filename']] = ['filename'=> $fn, 'contents' => LoadResourceFile($fn)];
        }

    }
    return $AllContent;
}


