' =========================================================
' Tittle: Netstat
' Version :
' Date: 30/07/2009
' Author: K Grey
' Company :
' Comment: Checks if VNC is running on port 5900, exists with error code 1 if true
' Version History  :
' =========================================================

Dim RE, Line, sh, shExec

set sh = CreateObject("Wscript.Shell") 
set shExec = sh.Exec("netstat -p tcp -an") 

Set RE = New RegExp
RE.Global = true
RE.IgnoreCase = true
RE.Pattern = "5900" 'vnc default port

Set RE1 = New RegExp
RE1.Global = true
RE1.IgnoreCase = true
RE1.Pattern = "established"

Do While Not shExec.StdOut.AtEndOfStream 
       Line = shExec.StdOut.ReadLine() 
       If RE.Test(Line) And RE1.Test(Line) Then 
       			'Debug.WriteLine Line 
       			WScript.Quit(1)
       End if
Loop 

WScript.Quit(0)