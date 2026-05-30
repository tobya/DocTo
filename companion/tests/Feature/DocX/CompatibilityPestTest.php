<?php

    it ('can read docx files',function (){

        $inputfiledir = 'inputfilesxls'. uniqid();
        $outputfiledir = 'outputfilesxls' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles(collect(['docx']), $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-WD')
            ->add('-f', $testinputfilesdir_temp  )
            ->add('-o', $testoutputdir_temp )
            ->add('-t', 'wdformatPDF')
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output->output());
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(1);
    });



    it ('can convert doc to docx with compatibility',function (){

        $inputfiledir = 'inputfiles_docx'. uniqid();
        $outputfiledir = 'outputfiles_docx' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles('plain', $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-WD')
            ->add('-f', $testinputfilesdir_temp  )
            ->add('-o', $testoutputdir_temp )
            ->add('-ox', '.docx')
            ->add('-t', 'wdFormatDocumentDefault')
            ->add('--COMPATIBILITY', '65535')
            ->add('-L',10)
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output->output());
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(5);

    // ensure -ox parameter is used.
    $file1 = $outputDirFiles->first();
    expect(str($file1)->endsWith('.docx'))->toBeTrue();


    });

    it ('can convert doc to docx with -c compatibility',function (){

        $inputfiledir = 'inputfiles_docx'. uniqid();
        $outputfiledir = 'outputfiles_docx' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles('plain', $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-WD')
            ->add('-f', $testinputfilesdir_temp  )
            ->add('-o', $testoutputdir_temp )
            ->add('-ox', '.docx')
            ->add('-t', 'wdFormatDocumentDefault')
            ->add('-c', '65535')
            ->add('-L',10)
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output->output());
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(5);

    // ensure -ox parameter is used.
    $file1 = $outputDirFiles->first();
    expect(str($file1)->endsWith('.docx'))->toBeTrue();


    });


    it ('can convert doc to alternative compatibility',function ($marker,$compatibility) {

        $inputfiledir = 'inputfiles_docx'. uniqid();
        $outputfiledir = 'outputfiles_docx' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles('plain', $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-WD')
            ->add('-f', $testinputfilesdir_temp  )
            ->add('-o', $testoutputdir_temp )
            ->add('-t', 'wdFormatDocumentDefault')
            ->add('--COMPATIBILITY', $compatibility)
            ->add('-L',10)
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output->output());
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(5);


    })->with(
        [
        ['wdCurrent',  65535],
        ['wdWord2003',  11],
        ['wdWord2007',  12],
        ['wdWord2010',  14],
        ['wdWord2013',  15],
    ]);
