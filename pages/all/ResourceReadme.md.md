Resource File Readme
==

There are several resource files present in DocTo.  Some are very straightforward, such as the help resource, which just needs to be kept up to date as items are added to the command line parameters.  Others however need a little bit of explanation.

WDConst.txt and XLConst.txt
--

Both of these files were extracted from firstly importing the main Word_TLB.pas via Delphis Import Type Library Feature.

**Then**


 - Copy Just the Constants section of the file to a new file.
 - Remove Type declarations (all these steps can be done easilty in a good text editor such as SublimeText).
 - Remove all Const declarations except for first one.
 - Turn into .pas file by adding interface keyword and end.
 - Save as Word_TLB_Constants.pas 
 - Take Const Section and save into new file WDConst.txt in res directory.
 - Remove all white space between declarations and all white space around `=`
 - Should be nothing but a Key=value file
 - Do the same for Excel.