<?php

  namespace App\Services;

  use Illuminate\Support\Collection;
  use Illuminate\Support\Facades\Storage;

  class FileGatherService
  {
        public static function GatherFiles(Collection $list, $tempDirName)
        {
            // remove exisitn files
            if (\Illuminate\Support\Facades\Storage::exists($tempDirName)){
                \Illuminate\Support\Facades\Storage::deleteDirectory($tempDirName);
            }

            $tempDirPath = Storage::path($tempDirName);
            $list->each(function ($dir) use ($tempDirName, $tempDirPath){
                $inputfilesdir = \Illuminate\Support\Facades\Storage::path('inputfiles\\' . $dir );
                $cmd = "xcopy \"$inputfilesdir\" \"$tempDirPath\\\"  ";
                $result =  \Illuminate\Support\Facades\Process::run( $cmd );
               // echo "\n $cmd \n";
               // echo "\n" . $result->output() . "\n";
            });
            return collect(Storage::listContents($tempDirName));
        }
  }
