{
    "title" : "{{$Command->Title}}" 
}

{{$Command->Title}}         
-

Sometimes you need to convert Word Documents that match a particular file pattern to {{$Command->FileTypeExt}}s.  
You can do this easily on the command line using [Docto](https://github.com/tobya/docto). 

The command line below shows how you can convert all the Microsoft Word Documents in a folder that match an input
filter to a {{$Command->FileTypeDescription}} file - {{$Command->FileTypeExt}}.  

Command Line 
-

 ````
 docto -WD -f 'c:\path\to\documents' -o 'c:\path\to\Outputfiles' --inputfilter='Accounting*.doc' -t {{$Command->FileFormat}}
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\to\documents' 
        -o 'c:\path\to\Outputfiles' 
        --inputfilter='Accounting*.doc'
        -t {{$Command->FileFormat}}
 ````
This will match all files in the directory c:\path\to\documents that match Accounting*.doc such as Accounting123.doc or 
Accounting_ClientA.doc

Command Line Explained 
-

 - `{{$Params->appwd->cmd}}` :  {{$Params->appwd->desc}}
 - `{{$Params->dashf->cmd}}` :  {{$Params->dashf->desc}} 
 - `{{$Params->dasho->cmd}}` :  {{$Params->dasho->desc}}
 - `{{$Params->ddinputfilter->cmd}}` :  {{$Params->ddinputfilter->desc}}
 - `{{$Params->dasht->cmd}}` :  {{$Params->dasht->desc}}




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a {{$Command->FileTypeExt}} file](ConvertDocToFile{{$Command->FileTypeExt}}.md);
@include('RelatedLinks') 

