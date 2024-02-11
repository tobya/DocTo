<?php

    it('test deletes files from directory', function (){
        if (\Illuminate\Support\Facades\Storage::exists('inputfilestemp')){

        \Illuminate\Support\Facades\Storage::deleteDirectory('inputfilestemp');
        }
       $testinputfilesdir = \Illuminate\Support\Facades\Storage::path('inputfiles');
       $testinputfilesdir_temp = \Illuminate\Support\Facades\Storage::path('inputfilestemp');
$dirfiles = \Illuminate\Support\Facades\Storage::listContents('inputfilestemp');

  $docfiles = $dirfiles->filter(function ($item){
    return str($item)->endsWith('.doc');
});
  print_r($docfiles->toArray());
$docfilecount = count($docfiles->toArray());
$dirfilescount = count(\Illuminate\Support\Facades\Storage::listContents('inputfilestemp')->toArray());
       $testoutputdir_temp = \Illuminate\Support\Facades\Storage::path('outputtemp2');
       \Illuminate\Support\Facades\Storage::createDirectory('outputtemp2');
       $cmd = "xcopy \"$testinputfilesdir\" \"$testinputfilesdir_temp\\\"  ";
       echo $cmd;
       $result =  \Illuminate\Support\Facades\Process::run( $cmd );
       echo $result->output();
        echo "\n $dirfilescount $docfilecount";
       $doctocmd = <<<CMD
..\\exe\\docto.exe -WD -f $testinputfilesdir_temp -fx .doc -o $testoutputdir_temp -t wdFormatPDF -R true
CMD;
       echo $doctocmd;
       $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       echo $output->output();
       expect(count(\Illuminate\Support\Facades\Storage::listContents('inputfilestemp')->toArray()))->tobe($dirfilescount - $docfilecount);

    });
