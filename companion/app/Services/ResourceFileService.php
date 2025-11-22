<?php

  namespace App\Services;

  class ResourceFileService
  {
    public static function LoadResourceFile($fn) {


        $fullfn = docto_path('\\src\\res\\' . $fn);
        return file_get_contents($fullfn);

    }

   public static  function LoadResourceFiles(){
        $allfiles = scandir(docto_path('\\src\\res\\'));

        foreach ($allfiles as $key => $fn) {

            if (strlen($fn) > 2 ){ // ignore . ..
                    //echo $fileinfo['filename'] . ':' . strlen($fileinfo['filename']);
                if (strpos( $fn , '__history') !== false){continue;}
                if (strpos( $fn , '__recovery') !== false){continue;}
                $info = pathinfo($fn);

                $AllContent[$info['filename']] = ['filename'=> $fn, 'contents' => static::LoadResourceFile($fn)];
            }

        }
        return $AllContent;
    }
  }
