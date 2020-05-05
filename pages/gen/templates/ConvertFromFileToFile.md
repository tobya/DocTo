{[$Command.Title]}        
==

{[$Command.Description]}

Command Line 
-

    docto {[foreach from=$Command.CMDValues item=Param]}{[$Param]} {[/foreach]}

Command Line Explained 
-

{[foreach from=$Command.CMDValues item=Param]}
 - {[$Param]} 
{[/foreach]}


Some other interesting commands
-

{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    


