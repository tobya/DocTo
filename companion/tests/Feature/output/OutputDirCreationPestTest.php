<?php

  use Illuminate\Support\Facades\Process;

  test('can create non existant directory', function () {
    $gatherdir = uniqid();
    $outputdir = uniqid();
    $files = \App\Services\FileGatherService::GatherFiles(collect(['single']),$gatherdir);
    $docto = config('services.docto.path');
    $inputdir = \Illuminate\Support\Facades\Storage::path($gatherdir);
    $outputdir = \Illuminate\Support\Facades\Storage::path($outputdir);
    $cmd = "$docto -WD -f $inputdir -o $outputdir -t wdFormatHTML";
   echo $cmd;
   $output = Process::run($cmd);
    expect(\Illuminate\Support\Facades\Storage::exists($outputdir))->toBeTrue();

});
