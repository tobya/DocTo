<?php



    it ('checks correct version',function (){

        $inputfiledir = 'inputfiles_docx'. uniqid();
        $outputfiledir = 'outputfiles_docx' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles('plain', $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-h')
            ->build();

        $result = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output->output());
        $outputString = $result->output();

        expect($outputString)->toContain('DocTo Version: 1.16.45');


    });

