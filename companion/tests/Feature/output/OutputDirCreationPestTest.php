<?php

  use Illuminate\Support\Facades\Process;

  test('can create non existant directory', function () {
    $gatherdir = uniqid();
    $outputdir = uniqid();
    $files = \App\Services\FileGatherService::GatherFiles(collect(['single']),$gatherdir);
    $docto = config('services.docto.path');
    $inputdir = \Illuminate\Support\Facades\Storage::path($gatherdir);
    $fulloutputdir = \Illuminate\Support\Facades\Storage::path($outputdir);
    $cmd = "$docto -WD -f $inputdir -o $fulloutputdir -t wdFormatHTML";
  // echo $cmd, $fulloutputdir;
   $output = Process::run($cmd);
    expect(\Illuminate\Support\Facades\Storage::exists($outputdir))->toBeTrue();

});
