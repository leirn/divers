' Script to test the performance of DNS calls
' Work on Microsot Windows
' The script forces to run with cscript instead of wscript
' 
' MIT License
' Copyright 2019 Laurent Vromman <laurent@vromman.org>
'
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Sub forceCScriptExecution
    Dim Arg, Str
    If Not LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
            If InStr( Arg, " " ) Then Arg = """" & Arg & """"
            Str = Str & " " & Arg
        Next
        ret = CreateObject( "WScript.Shell" ).Run ("cmd /k cscript //nologo """ & WScript.ScriptFullName & """ " & Str,1,true)
        WScript.Quit ret
    End If
End Sub
forceCScriptExecution


Set re = New RegExp
 With re
	.Pattern    = "^\s+([a-z0-9]+(-[a-z0-9]+)*\.+[a-z]{2,}){1}"
	.IgnoreCase = True
	.Global     = False
 End With

Set objShell = CreateObject("WScript.Shell")
comspec = objShell.ExpandEnvironmentStrings("%comspec%")

Set objExec = objShell.Exec(comspec & " /c ipconfig /displaydns")

tmp = ""
URLs = Array()
i = 0

Do While Not objExec.StdOut.AtEndOfStream
	tmp = objExec.StdOut.ReadLine()

	If re.Test(tmp) Then
		REDIM PRESERVE URLs(i + 1)
		URLs(i) = Trim(tmp)
		i = i + 1
	End If
Loop

objShell.Run comspec & " /c nslookup /flushdns", 0, True

For j = 0 to i - 1
	'WScript.Echo comspec & " /c nslookup " & URLs(j)
	dtmStartTime = Timer
    objShell.Run comspec & " /c nslookup " & URLs(j), 0, True
	executionTime = Timer - dtmStartTime
	WScript.Echo("Time to get " & URLs(j) & ": " & FormatNumber(executionTime, 2) & " seconds.")
Next

Set objShell = Nothing
Set objExec = Nothing