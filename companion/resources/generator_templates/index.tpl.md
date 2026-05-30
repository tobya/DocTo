DocTo Examples
==

Here is a list of examples on how to use DocTo.  You should find details of what you would like to do somewhere here.

@foreach( $Commands as $CommandType)

{{$CommandType->Description}}
==

@foreach ( $CommandType->Items as $C)
 - [{{$C->fntitle}}]({{$C->fn}})
@endforeach

@endforeach

Help Text
--

DocTo outputs a [help file](HelpLog.md) which provides details of all options. 

Generated {{now()}}
