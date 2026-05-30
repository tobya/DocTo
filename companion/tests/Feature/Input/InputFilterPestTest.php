<?php

    it('can filter input files', function (){
        $inputfiledir = 'inputfiles'. uniqid();
        $outputfiledir = 'outputfiles_docz' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles('mixed', $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-WD')
            ->add('-f', $testinputfilesdir_temp  )
            ->add('--inputfilter', '*pie*.doc'   ) // should match 5
            ->add('-o', $testoutputdir_temp )
            ->add('-t', 'wdFormatPDF')
            ->add('-L',10)
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
      //  print_r($output->output());
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(5);

    // ensure -ox parameter is used.
    $file1 = $outputDirFiles->first();
    expect(str($file1)->endsWith('.pdf'))->toBeTrue();


    });



    it('can use multi input files', function (){
        $inputfiledir = 'inputfiles'. uniqid();
        $outputfiledir = 'outputfiles_docz' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles('mixed', $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-WD')
            ->add('-f', $testinputfilesdir_temp  )
            ->add('--inputfilter', '*pie*'   ) // should match 11
            ->add('-o', $testoutputdir_temp )
            ->add('-t', 'wdFormatPDF')
            ->add('-L',10)
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
      //  print_r($output->output());
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(11);

    // ensure -ox parameter is used.
    $file1 = $outputDirFiles->first();
    expect(str($file1)->endsWith('.pdf'))->toBeTrue();


    });



    it('can use non doc input files', function ($filter,$count){
        $inputfiledir = 'inputfiles'. uniqid();
        $outputfiledir = 'outputfiles_docz' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles('mixed', $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-WD')
            ->add('-f', $testinputfilesdir_temp  )
            ->add('--INPUTFILEEXTENSION', $filter   )
            ->add('-o', $testoutputdir_temp )
            ->add('-t', 'wdFormatPDF')
            ->add('-L',10)
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output->output());
        expect( $output->exitCode())->tobe(0);
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe($count);

    // ensure -ox parameter is used.
    $file1 = $outputDirFiles->first();
    expect(str($file1)->endsWith('.pdf'))->toBeTrue();


    })->with(
        [
            ['.txt',2],
            ['.rtf',2],
            ['.odt',2],
            ['.docx',1],

        ]
    );
