<?php

it('test deletes files from directory', function (){
        // setup
        if (\Illuminate\Support\Facades\Storage::exists('inputfilestemp')){
            \Illuminate\Support\Facades\Storage::deleteDirectory('inputfilestemp');
        }
       $testinputfilesdir = \Illuminate\Support\Facades\Storage::path('inputfiles\\plain');
       $testinputfilesdir_temp = \Illuminate\Support\Facades\Storage::path('inputfilestemp');



       $testoutputdir_temp = \Illuminate\Support\Facades\Storage::path('outputtemp2');
      // echo "\n". $testoutputdir_temp;
       \Illuminate\Support\Facades\Storage::createDirectory('outputtemp2');
       $cmd = "xcopy \"$testinputfilesdir\" \"$testinputfilesdir_temp\\\"  ";
      // echo "\n". $cmd;
       $result =  \Illuminate\Support\Facades\Process::run( $cmd );

            //echo "\n" . $result->output() . "\n";
           $dirfiles = collect(\Illuminate\Support\Facades\Storage::listContents('inputfilestemp'));
          $docfiles = $dirfiles->filter(function ($item){
            return str($item->path())->endsWith('.doc');
        });

         expect($docfiles->count())->toBeGreaterThan(0);
            $docfilecount = count($docfiles->toArray());

        $dirfilescount = $dirfiles->count();
        // do conversion
       $doctocmd = <<<CMD
..\\src\\docto.exe -WD -f $testinputfilesdir_temp -fx .doc -o $testoutputdir_temp -t wdFormatPDF -R true
CMD;
      // echo $doctocmd;
       $output = \Illuminate\Support\Facades\Process::run($doctocmd);
     //  echo $output->output();

    // check files have been converted and originoals have been created.
       expect(collect(\Illuminate\Support\Facades\Storage::listContents('inputfilestemp'))->count())->tobe($dirfilescount - $docfilecount);
       expect(collect(\Illuminate\Support\Facades\Storage::listContents('outputtemp2'))->count())->tobe( $docfilecount);


    });
