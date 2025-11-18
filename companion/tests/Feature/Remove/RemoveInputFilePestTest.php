<?php

it('test deletes files from directory', function (){
        // setup
      // $testinputfilesdir = \Illuminate\Support\Facades\Storage::path('inputfiles\\plain');
       $testinputfilesdir_temp = \Illuminate\Support\Facades\Storage::path('inputfilestemp');


       $testinputfilesdir_temp = \Illuminate\Support\Facades\Storage::path('inputfilestemp');

       $testoutputdir_temp = \Illuminate\Support\Facades\Storage::path('outputtemp2');

       \Illuminate\Support\Facades\Storage::createDirectory('outputtemp2');

     //  $cmd = "xcopy \"$testinputfilesdir\" \"$testinputfilesdir_temp\\\"  ";
      // echo "\n". $cmd;
    //   $result =  \Illuminate\Support\Facades\Process::run( $cmd );

            //echo "\n" . $result->output() . "\n";
         //  $dirfiles = collect(\Illuminate\Support\Facades\Storage::listContents('inputfilestemp'));
    $dirfiles = \App\Services\FileGatherService::GatherFiles(collect(['plain']),'inputfilestemp');
          $docfiles = $dirfiles->filter(function ($item){
            return str($item->path())->endsWith('.doc');
        });

         expect($docfiles->count())->toBeGreaterThan(0);
            $docfilecount = count($docfiles->toArray());

        $dirfilescount = $dirfiles->count();
        // do conversion
        $docto = config('services.docto.path');
      // $doctocmd = "$docto -WD -f $testinputfilesdir_temp -fx .doc -o $testoutputdir_temp -t wdFormatPDF -R true";
       $doctocmd = \App\Services\DocToCommandBuilder::docto()
           ->add('-WD')
           ->add('-f',$testinputfilesdir_temp)
           ->add('-fx','.doc')
           ->add('-o',$testoutputdir_temp)
           ->add('-t wdFormatPDF')
           ->add('-R','true') // the important one fo rthis test
           ->build();
      // echo $doctocmd;
       $output = \Illuminate\Support\Facades\Process::run($doctocmd);


    // check files have been converted and originoals have been created.
       expect(collect(\Illuminate\Support\Facades\Storage::listContents('inputfilestemp'))->count())->tobe($dirfilescount - $docfilecount);
       expect(collect(\Illuminate\Support\Facades\Storage::listContents('outputtemp2'))->count())->tobe( $docfilecount);


    });



it('doesnt delete files from directory', function (){
        // setup
      // $testinputfilesdir = \Illuminate\Support\Facades\Storage::path('inputfiles\\plain');
       $testinputfilesdir_temp = \Illuminate\Support\Facades\Storage::path('inputfilestemp');


       $testinputfilesdir_temp = \Illuminate\Support\Facades\Storage::path('inputfilestemp');

       $testoutputdir_temp = \Illuminate\Support\Facades\Storage::path('outputtemp2');

       \Illuminate\Support\Facades\Storage::createDirectory('outputtemp2');

     //  $cmd = "xcopy \"$testinputfilesdir\" \"$testinputfilesdir_temp\\\"  ";
      // echo "\n". $cmd;
    //   $result =  \Illuminate\Support\Facades\Process::run( $cmd );

            //echo "\n" . $result->output() . "\n";
         //  $dirfiles = collect(\Illuminate\Support\Facades\Storage::listContents('inputfilestemp'));
    $dirfiles = \App\Services\FileGatherService::GatherFiles(collect(['plain']),'inputfilestemp');
          $docfiles = $dirfiles->filter(function ($item){
            return str($item->path())->endsWith('.doc');
        });

         expect($docfiles->count())->toBeGreaterThan(0);
            $docfilecount = count($docfiles->toArray());

        $dirfilescount = $dirfiles->count();
        // do conversion
        $docto = config('services.docto.path');
      // $doctocmd = "$docto -WD -f $testinputfilesdir_temp -fx .doc -o $testoutputdir_temp -t wdFormatPDF -R true";
       $doctocmd = \App\Services\DocToCommandBuilder::docto()
           ->add('-WD')
           ->add('-f',$testinputfilesdir_temp)
           ->add('-fx','.doc')
           ->add('-o',$testoutputdir_temp)
           ->add('-t wdFormatPDF')
         //  ->add('-R','true') // the important one fo rthis test
           ->build();
      // echo $doctocmd;
       $output = \Illuminate\Support\Facades\Process::run($doctocmd);


    // check files have been converted and originoals have been created.
       expect(collect(\Illuminate\Support\Facades\Storage::listContents('inputfilestemp'))->count())->tobe( $docfilecount);

       expect(collect(\Illuminate\Support\Facades\Storage::listContents('outputtemp2'))->count())->tobe( $docfilecount);


    });
