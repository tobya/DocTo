<?php

    /**
     * This set of tests relies on a xls file with 4 tabs
     * Sheet1
     * Sheet2
     * Tab3
     * Sheet4
     *
     * Sheet4 is blank.
     */


    it('can get named sheet in a multi sheet xls', function () {

        $inputfiledir = 'inputfilesxls';
        $outputfiledir = 'outputfilesxls' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);
 $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));
    expect($outputDirFiles->count())->tobe(0);
        $dirfiles = \App\Services\FileGatherService::GatherFiles(collect(['multisheet']), $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-XL')
            ->add('-f', $testinputfilesdir_temp .'\\Book1 MultiSheet Test.xlsx' )
            ->add('-o', $testoutputdir_temp .'\\Book1 MultiSheet Test.pdf')
            ->add('-t', 'xlPDF')
            ->add('--sheets', 'Tab3')
            ->add('-L 10')
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output);
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(1);
        // Sheet named
        $sheetNamed = $outputDirFiles->filter(function ($item) {
           // echo $item . "\n";
           if (str($item)->contains('Book1 MultiSheet Test_(Sheet1).pdf') ) return true;
           if (str($item)->contains('Book1 MultiSheet Test_(Sheet2).pdf') ) return true;
           Return false;
        });


        $sheetNamed3 = $outputDirFiles->filter(function ($item) {
           // echo $item . "\n";
           if (str($item)->contains('Book1 MultiSheet Test_(Tab3).pdf') ) return true;

        });

        expect($sheetNamed->count())->tobe(0);
        expect($sheetNamed3->count())->tobe(1);

      //  Storage::deleteDirectory($outputfiledir);
    });


        it('can pdf each sheet in a multi sheet xls', function () {

        $inputfiledir = 'inputfilesxls';
        $outputfiledir = 'outputfilesxls' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles(collect(['multisheet']), $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-XL')
            ->add('-f', $testinputfilesdir_temp .'\\Book1 MultiSheet Test.xlsx' )
            ->add('-o', $testoutputdir_temp .'\\Book1 MultiSheet Test.pdf')
            ->add('-t', 'xlPDF')
            ->add('--sheets', '1,2')
            ->add('-L 10')
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output);
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(2);
        // Sheet named
        $sheetNamed = $outputDirFiles->filter(function ($item) {
           // echo $item . "\n";
           if (str($item)->contains('Book1 MultiSheet Test_(Sheet1).pdf') ) return true;
           if (str($item)->contains('Book1 MultiSheet Test_(Sheet2).pdf') ) return true;
           Return false;
        });


        $sheetNamed3 = $outputDirFiles->filter(function ($item) {
           // echo $item . "\n";
           if (str($item)->contains('Book1 MultiSheet Test_(Tab3).pdf') ) return true;

        });

        expect($sheetNamed->count())->tobe(2);
        expect($sheetNamed3->count())->tobe(0);

        Storage::deleteDirectory('outputtempxls2');
    });


        it('can pdf output all sheets multi sheet xls', function () {

        $inputfiledir = 'inputfilesxls';
        $outputfiledir = 'outputfilesxls' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles(collect(['multisheet']), $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-XL')
            ->add('-f', $testinputfilesdir_temp .'\\Book1 MultiSheet Test.xlsx' )
            ->add('-o', $testoutputdir_temp .'\\Book1 MultiSheet Test.pdf')
            ->add('-t', 'xlPDF')
            ->add('--allsheets')
            ->add('-L 10')
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output);
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(3);
        // Sheet named
        $sheetNamed = $outputDirFiles->filter(function ($item) {
           // echo $item . "\n";
           if (str($item)->contains('Book1 MultiSheet Test_(Sheet1).pdf') ) return true;
           if (str($item)->contains('Book1 MultiSheet Test_(Sheet2).pdf') ) return true;
           if (str($item)->contains('Book1 MultiSheet Test_(Tab3).pdf') ) return true;
           // this is not expected to match.  Emtpy sheets dont output
           if (str($item)->contains('Book1 MultiSheet Test_(Sheet4).pdf') ) return true;
           Return false;
        });

        expect($sheetNamed->count())->tobe(3);


        Storage::deleteDirectory($outputfiledir);
    });

 it('outputs correct all sheets multi sheet xls', function ($format, $ext) {

        $inputfiledir = 'inputfilesxls';
        $outputfiledir = 'outputfilesxls' . uniqid();
        // setup
        $testinputfilesdir_temp = Storage::path($inputfiledir);
        $testoutputdir_temp = Storage::path($outputfiledir);

        Storage::createDirectory($outputfiledir);

        $dirfiles = \App\Services\FileGatherService::GatherFiles(collect(['multisheet']), $inputfiledir);

        $doctocmd = \App\Services\DocToCommandBuilder::docto()
            ->add('-XL')
            ->add('-f', $testinputfilesdir_temp .'\\Book1 MultiSheet Test.xlsx' )
            ->add('-o', $testoutputdir_temp .'\\TabTest.' . $ext)
            ->add('-t', $format)
            ->add('--allsheets')
            ->add('-L 10')
            ->build();

        $output = \Illuminate\Support\Facades\Process::run($doctocmd);
       // print_r($output->output());
        $outputDirFiles = collect(\Illuminate\Support\Facades\Storage::allFiles($outputfiledir));

    expect($outputDirFiles->count())->toBeGreaterThan(0);
    expect($outputDirFiles->count())->tobe(4);


        // Sheet named
        $fileText = file_get_contents(Storage::path($outputfiledir . '\\TabTest_(Tab3).' . $ext));
        expect(str($fileText)->contains('This is Tab3'))->toBeTrue();
        expect(str($fileText)->contains('This is Sheet 1'))->not()->toBeTrue();

        // Sheet named
        $fileText = file_get_contents(Storage::path($outputfiledir . '\\TabTest_(Sheet1).' . $ext));
        expect(str($fileText)->contains('This is Sheet 1'))->toBeTrue();
        expect(str($fileText)->contains('This is Tab3'))->not()->toBeTrue();



    })->with([
        ['xlCSV', 'csv'],
        ['xlUnicodeText', 'txt'],
        ['xlCSVWindows', 'csv'],
        ['xlTextWindows', 'txt'],
 ]);
