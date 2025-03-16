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

                print_r($CommandBlock);
               // $smarty->assign('CommandBlock' , $CommandBlock);
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
                 // echo "\n Title : $CommandBlock[Title] \n";
                 // print_r($CommandBlock);
                 // print_r($Item);
                 // echo "\n";

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
                echo "\n -----";
                print_r( $CommandBlock['Template']);
                echo " -----\n";

                $MDFile = Blade::render(file_get_contents(base_path('\\resources\\generator_templates\\' .  $CommandBlock['Template'])),[
                        'Params' => json_decode(json_encode($Params)),
                        'ResourceFiles' => $AllResourceFiles,
                        'Command' =>  json_decode(json_encode($Item)),
                        'CommandBlock' =>   json_decode(json_encode($CommandBlock)),
                    ]);
                echo "Create File : $Fn\n";
                file_put_contents(base_path('/storage/app/all/' . $Fn), $MDFile);



            }
        }

        die('done to here');
        //**************
        //  Create index file here
        //**********************

      //  $smarty->assign('GeneratedTime', date('H:i:s Ymd'));
     //   $smarty->Assign('Commands',$Commands);
       // $Indexmd = $smarty->fetch('index.tpl.md');
        $Indexmd = view('index.tpl.md');
        //print_r($Indexmd);
        file_put_contents('../all/index.md', $Indexmd);
        echo "Create File : index.md \n";

        $HelpFile = LoadResourceFile('HelpLog.txt');


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
            file_put_contents('../all/' . $fileinfo['filename'] . '.md', $Contents);
            echo "\nCreate Resource File : " . $fileinfo['filename'] . '.md';
        }
    }
}
