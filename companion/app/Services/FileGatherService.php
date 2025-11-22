<?php

  namespace App\Services;

  use Illuminate\Support\Collection;
  use Illuminate\Support\Facades\Storage;

  class FileGatherService
  {
      /**
       * Gather files will take files from specific subdir of the resource InputFiles directory and
       * copy them to a temporary directory wehre they can be used for a test.
       * @param Collection $list
       * @param $tempDirName
       * @return Collection
       * @throws \League\Flysystem\FilesystemException
       */
        public static function GatherFiles(Collection | string $list, $tempDirName)
        {
            if (is_string($list)){
                $list = collect([$list]);
            }

            // remove exisitn files
            if (\Illuminate\Support\Facades\Storage::exists($tempDirName)){
                \Illuminate\Support\Facades\Storage::deleteDirectory($tempDirName);
            }

            $tempDirPath = Storage::path($tempDirName);
            $list->each(function ($inputdir) use ($tempDirName, $tempDirPath){
                $inputfilesdir = \Illuminate\Support\Facades\Storage::disk('inputfiles')->path($inputdir );
                $cmd = "xcopy \"$inputfilesdir\" \"$tempDirPath\\\"  ";
                $result =  \Illuminate\Support\Facades\Process::run( $cmd );
            });
            return collect(Storage::listContents($tempDirName));
        }
  }
