<?php

  use Illuminate\Support\Facades\Process;

  test('can create non existant directory', function () {
    $gatherdir = uniqid();
    $outputdir = uniqid();
    $files = \App\Services\FileGatherService::GatherFiles(collect(['single']),$gatherdir);
    $docto = config('services.docto.path');
    $inputdir = \Illuminate\Support\Facades\Storage::path($gatherdir);
    $fulloutputdir = \Illuminate\Support\Facades\Storage::path($outputdir);
   // $cmd = "$docto -WD -f $inputdir -o $fulloutputdir -t wdFormatHTML";
    $cmd = \App\Services\DocToCommandBuilder::docto()->add('-WD')->add('-f',$inputdir)
        ->add('-o',$fulloutputdir)
        ->add('-t wdFormatHTML')->build();
    \Illuminate\Support\Facades\Log::debug($cmd);
    \Illuminate\Support\Facades\Log::debug($fulloutputdir);
   $output = Process::run($cmd);
    expect(\Illuminate\Support\Facades\Storage::exists($outputdir))->toBeTrue();

});

  test('will not create extension directory', function () {
    $gatherdir = uniqid();
    $outputdir = uniqid();
    $files = \App\Services\FileGatherService::GatherFiles(collect(['single']),$gatherdir);
    $docto = config('services.docto.path');
    $inputdir = \Illuminate\Support\Facades\Storage::path($gatherdir);
    $fulloutputdir = \Illuminate\Support\Facades\Storage::path($outputdir);
    // a bug caused docto to create and ouput to extension directory.
    $pdfdir = $fulloutputdir . '\\.pdf';
   // $cmd = "$docto -WD -f $inputdir -o $fulloutputdir -t wdFormatPDF";
    $cmd = \App\Services\DocToCommandBuilder::docto()->add('-WD')
        ->add('-f',$inputdir)
        ->add('-o',$fulloutputdir)
        ->add('-t wdFormatPDF')->build();
    \Illuminate\Support\Facades\Log::debug($cmd);
    \Illuminate\Support\Facades\Log::debug($fulloutputdir);
   $output = Process::run($cmd);
    expect(\Illuminate\Support\Facades\Storage::exists($pdfdir))->toBeFalse();

});




  test('can create non existant directory with bulder', function () {
    $gatherdir = uniqid();
    $outputdir = uniqid();
    $files = \App\Services\FileGatherService::GatherFiles(collect(['single']),$gatherdir);
    $docto = config('services.docto.path');
    $inputdir = \Illuminate\Support\Facades\Storage::path($gatherdir);
    $fulloutputdir = \Illuminate\Support\Facades\Storage::path($outputdir);
    $cmd = \App\Services\DocToCommandBuilder::docto()->add('-WD')
        ->add('-f',$inputdir)
        ->add('-o',$fulloutputdir)
        ->add('-t wdFormatHTML')->build();
   // $cmd = "$docto -WD -f $inputdir -o $fulloutputdir -t wdFormatHTML";
    \Illuminate\Support\Facades\Log::debug($cmd);
    \Illuminate\Support\Facades\Log::debug($fulloutputdir);
   $output = Process::run($cmd);
    expect(\Illuminate\Support\Facades\Storage::exists($outputdir))->toBeTrue();

});
