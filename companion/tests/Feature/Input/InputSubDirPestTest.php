<?php


it('can get all files in subdir', function (){
   $inputfiledir = 'inputfiles_sub_'. uniqid();
        $outputfiledir = 'outputfiles_subdir' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles('plain', $inputfiledir);
        $subdirfiles = \App\Services\FileGatherService::GatherFiles('plain', $inputfiledir . '\\subdir');

        $allfilestotal = $dirfiles->count() + $subdirfiles->count();

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-WD')
            ->add('-f', $testinputfilesdir_temp  )
            ->add('-o', $testoutputdir_temp )
            ->add('-t', 'wdFormatPDF')
            ->add('-L',10)
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
      //  print_r($output->output());
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));


    expect($outputDirFiles->count())->tobe($allfilestotal);

    // ensure -ox parameter is used.
    $file1 = $outputDirFiles->first();
    expect(str($file1)->endsWith('.pdf'))->toBeTrue();

});

it('can get all files in base dir but not subdir', function ($command){
   $inputfiledir = 'inputfiles'. uniqid();
        $outputfiledir = 'outputfiles_docz' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles('plain', $inputfiledir);
        $subdirfiles = \App\Services\FileGatherService::GatherFiles('plain', $inputfiledir . '\\subdir');

        $allfilestotal = $dirfiles->count() + $subdirfiles->count();

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-WD')
            ->add('-f', $testinputfilesdir_temp  )
            ->add('-o', $testoutputdir_temp )
            ->add('-t', 'wdFormatText')
            ->add($command) // should not load files from /subdir
            ->add('-L',10)
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
      //  print_r($output->output());
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));


    expect($dirfiles->count())->toBe(5);
    expect($outputDirFiles->count())->toBe($dirfiles->count());
    expect($outputDirFiles->count())->toBeLessThan($allfilestotal);

    // ensure -ox parameter is used.
    $file1 = $outputDirFiles->first();
    expect(str($file1)->endsWith('.txt'))->toBeTrue();

})->with([
    ['--NO-RECURSE'],
    ['--NO-SUBDIR'],
    ['--NO-SUBDIRS'],

]);
