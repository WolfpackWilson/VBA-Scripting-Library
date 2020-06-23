# VBA Scripting Library
The goal of this repository is to recreate some of the scripting library included
on Windows OS so that programs can work on Mac OS, too.

## History
When running programs on the Mac OS, an error appears:<br>
`Error: Run-time error ’429’ ActiveX component can’t create object`

It appears that the 
[Mac OS doesn't have the Scripting Runtime Library](https://stackoverflow.com/questions/4670420/how-can-i-install-use-scripting-filesystemobject-in-excel-2010-for-mac). 
As such, anytime `CreateObject("Scripting.<object>")` is used, this error will appear on Mac OS.
