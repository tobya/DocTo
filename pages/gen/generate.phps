<?php

require "config.php";

require "loadcommands.php";


$smarty->assign('Params',$Explain);

$AllResourceFiles = LoadResourceFiles();
$smarty->assign('ResourceFiles',$AllResourceFiles);

foreach ($Commands as $CommandName => $CommandBlock) {
    
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
        
        $smarty->assignByRef('Command', $Item);
        if (isset($CommandBlock['Title'])){

            $title = $smarty->fetch('string:'.$CommandBlock['Title']);
                        
        } else {
            $title = $Fn;
        }

        $Item['Title'] = $title;

        // Store fn for linking to in index page.
        $Commands[$CommandName]['Items'][$keytag]['fn'] = $Fn;
        $Commands[$CommandName]['Items'][$keytag]['fntitle'] = $title;
        

        $MDFile = $smarty->fetch($CommandBlock['Template']);
        
        echo "Create File : $Fn\n";
        file_put_contents('../all/' . $Fn, $MDFile);



    }
}

//**************
//  Create index file here
//**********************

$smarty->assign('GeneratedTime', date('H:i:s Ymd'));
$smarty->Assign('Commands',$Commands);
$Indexmd = $smarty->fetch('index.tpl.md');
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

