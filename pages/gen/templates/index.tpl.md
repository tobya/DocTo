DocTo Examples
==

Here is a list of examples on how to use DocTo.  You should find details of what you would like to do somewhere here.

{[foreach from=$Commands item=CommandType]}

{[$CommandType.Description]}
==

{[foreach from=$CommandType.Items item=C]}
 - [{[$C.fn]}]({[$C.fn]})
{[/foreach]}

{[/foreach]}