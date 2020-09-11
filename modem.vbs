' =========================================================
' Tittle: Modem
' Version : v0.1
' Date: 01/10/2000
' Author: K Grey
' Comapny : None
' Comment: simple connection to phone using com port and 'AT' commad set

' Version History  :
' =========================================================

' List Modem Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_POTSModem")

For Each objItem in colItems
    Wscript.Echo "Attached To: " & objItem.AttachedTo
    Wscript.Echo "Model: " & objItem.Model
	WScript.Echo
Next

