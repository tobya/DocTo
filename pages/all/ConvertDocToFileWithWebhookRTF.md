{
    "title" : "How do I Convert a Microsoft Word Doc to a Windows Rich Text Format and fire a webhook on conversion? " 
}

Windows Rich Text Format 
==

How do I Convert a Microsoft Word Doc to a Windows Rich Text Format (RTF) and fire a webhook on conversion?         
-

It is very simple to convert a Word Document to a RTF file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Word, but sometimes it helps to be able to do it from the command line.  And its really simple to fire a webhook that you can recieve at any website.  Just use the `-W` parameter.   

The command line below shows how you can convert a Microsoft Word Document to a Windows Rich Text Format file - RTF and fire a webhook.

Command Line 
-

 ````
 docto -WD -f 'c:\path\Document.doc' -o 'c:\path\Document.RTF' -t wdFormatRTF -W 'https://example.com/mywebhook.php'
 ````
 or easier to read
 ````
 docto -WD  -f 'c:\path\Document.doc' 
            -o 'c:\path\Document.RTF' 
            -t wdFormatRTF
            -W 'https://example.com/mywebhook.php'
 ````

Command Line Explained 
-

 - `-WD` -  This is a conversion using Microsoft Word.  This is not required but makes it easier to read
 - `-f` -  The File or directory to be converted 
 - `-o` -  The File or Directory where you would like the converted file to be written to.
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



Some other interesting commands
-

You might find some of the following commands also interesting.

    

