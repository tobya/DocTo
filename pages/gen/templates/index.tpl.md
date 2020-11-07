DocTo Examples
==

Here is a list of examples on how to use DocTo.  You should find details of what you would like to do somewhere here.

{[foreach from=$Commands item=CommandType]}

{[$CommandType.Description]}
==

{[foreach from=$CommandType.Items item=C]}
 - [{[$C.fntitle]}]({[$C.fn]})
{[/foreach]}

{[/foreach]}

Help Text
--

DocTo outputs a [help file](HelpLog.md) which provides details of all options. 

Generated {[$GeneratedTime]}