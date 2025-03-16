<?php

namespace App\Console\Commands;

use App\Services\CommandInfoService;
use App\Services\ResourceFileService;
use Illuminate\Console\Command;
use Illuminate\Support\Facades\Blade;

class GenerateDocumentationPages extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'docto:generate-pages';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Generates Documentation Pages with info on all the available commands.';

    /**
     * Execute the console command.
     */
    public function handle()
    {


        $Commands = CommandInfoService::Commands();
        $Params = CommandInfoService::Explanations();

        $AllResourceFiles = ResourceFileService::LoadResourceFiles();

        foreach ($Commands as $CommandName => $CommandBlock) {


            foreach ($CommandBlock['Items'] as $keytag => $Item) {
                # code...
                // for CommandBlocks that use a single format
                if (isset($CommandBlock['FileFormat']) ){
                    $Item['FileTypeExt'] = $CommandBlock['FileFormat']['FileTypeExt'];
                    $Item['FileTypeDescription'] = $CommandBlock['FileFormat']['FileTypeDescription'];

                    $Item['FileFormat'] = $CommandBlock['FileFormat']['FileFormat'];
                }



                if (!file_exists('../all/')){
                    mkdir('../all/');
                }

                $Fn =   $CommandName . @$Item['FileTypeTitleExtra'] . $Item['FileTypeExt'] . '.md' ;

                if (isset($CommandBlock['Title'])){


                    $title = Blade::render($CommandBlock['Title'],[
                        'Params' => json_decode(json_encode($Params)),
                        'ResourceFiles' => $AllResourceFiles,
                        'Command' => json_decode(json_encode( $Item)),
                        'CommandBlock' => json_decode(json_encode($CommandBlock)),
                    ]);
                  //  echo "\n Title : $title";
                } else {
                    $title = $Fn;
                }

                $Item['Title'] = $title;

                // Store fn for linking to in index page.
                $Commands[$CommandName]['Items'][$keytag]['fn'] = $Fn;
                $Commands[$CommandName]['Items'][$keytag]['fntitle'] = $title;


                $MDFile = Blade::render(file_get_contents(base_path('\\resources\\generator_templates\\' .  $CommandBlock['Template'])),[
                        'Params' => json_decode(json_encode($Params)),
                        'ResourceFiles' => $AllResourceFiles,
                        'Command' =>  json_decode(json_encode($Item)),
                        'CommandBlock' =>   json_decode(json_encode($CommandBlock)),
                    ]);
                echo "Create File : $Fn\n";
                file_put_contents(docto_path('/pages/all/' . $Fn), $MDFile);



            }
        }

        //**************
        //  Create index file here
        //**********************

        $Indexmd =  Blade::render(file_get_contents(base_path('\\resources\\generator_templates\\index.tpl.md')),[
                        'Params' => json_decode(json_encode($Params)),
                        'ResourceFiles' => $AllResourceFiles,
                        'Commands' =>  json_decode(json_encode($Commands)),

                        'CommandBlock' =>   json_decode(json_encode($CommandBlock)),
                    ]);

        file_put_contents(docto_path('/pages/all/index.md'), $Indexmd);

        echo "Create File : index.md \n";

        $HelpFile = ResourceFileService::LoadResourceFile('HelpLog.txt');


        // Wrap helpfile in ```` to ensure formatting.
        $HelpFile = "{
            \"title\" : \"Help Log File\"
        }

        ````
        $HelpFile
        ````";

        file_put_contents('../all/HelpLog.md', $HelpFile);
        echo "Create File : HelpLog.md \n";

        foreach ($AllResourceFiles as $key => $fileinfo) {



            $Contents = "


            $fileinfo[contents]

                ";
            $ResourceFiles[$fileinfo['filename']] = $Contents;
            file_put_contents(docto_path('/pages/all/' . $fileinfo['filename'] . '.md'), $Contents);
            echo "\nCreate Resource File : " . $fileinfo['filename'] . '.md';
        }
    }
}
