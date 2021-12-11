{
    "title" : "How do I Convert a Microsoft Word Doc to a Latest Microsoft Office Word 365 Version Format  and fire a webhook on conversion? " 
}

Latest Microsoft Office Word 365 Version Format  
==

How do I Convert a Microsoft Word Doc to a Latest Microsoft Office Word 365 Version Format  (DOC) and fire a webhook on conversion?         
-

It is very simple to convert a Word Document to a DOC file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Word, but sometimes it helps to be able to do it from the command line.  And its really simple to fire a webhook that you can recieve at any website.  Just use the `-W` parameter.   

The command line below shows how you can convert a Microsoft Word Document to a Latest Microsoft Office Word 365 Version Format  file - DOC and fire a webhook.

Command Line 
-

 ````
 docto -WD -f 'c:\path\Document.doc' -o 'c:\path\Document.DOC' -t wdFormatDocumentDefault -W 'https://example.com/mywebhook.php'
 ````
 or easier to read
 ````
 docto -WD  -f 'c:\path\Document.doc' 
            -o 'c:\path\Document.DOC' 
            -t wdFormatDocumentDefault
            -W 'https://example.com/mywebhook.php'
 ````

Command Line Explained 
-

 - `-WD` -  This is a conversion using Microsoft Word. 
 - `-f` -  The File or directory to be converted 
 - `-o` -  The Output File or Directory where you would like the converted file to be written to.
 - `-T` -  The file format type that is being converted to
 - `-W` -  Fire a webhook to the provided URL. See Below for more details


Web Hook Help
-

The Webhook URL will be called on the following events with the following parameters

  - File Conversion
    - action=convert
    - type=wdFormatType (or int if no matching format type)
    - ouputfilename=File being written to.
    - inputfilename=File being converted.

  - Error
    - action=error
    - type=wdFormatType (or int if no matching format type)
    - ouputfilename=File being written to.
    - inputfilename=File being converted.
    - error=Error Message

Return value is ignored, no errors are logged.  This is a fire and forget Webhook.

The URL will look like this

     https://example.com/webhooks/docto/webhook_test.php?action=convert&type=wdFormatPDF&outputfilename=D:%5CDevelopment%5CGitHub%5CDocTo%5Ctest%5CGeneratedFiles%5Cpie3.pdf&inputfilename=D:%5CDevelopment%5CGitHub%5CDocTo%5Ctest%5CInputfiles%5Cpie3.doc



Some other interesting commands
-

You might find some of the following commands also interesting.

    

