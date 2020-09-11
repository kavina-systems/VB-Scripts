' =========================================================
' Tittle: Phone_SMS
' Version : v0.1
' Date: 01/03/2001
' Author: K Grey
' Comapny : None
' Comment: simple connection to phone using com port and 'AT' commad set

' Version History  :
' =========================================================

Const ForWriting = 2, TriStateFalse = 0, DelayMilliSecs = 1000

Dim PhoneNum, Message

' Script needs two arguments. First one is the mobile phone to send message to
' Second one is the message to send (needs to be included in double quotes
' when calling the script).

'PhoneNum = objArgs(0)
'Message = objArgs(1)
PhoneNum = "1234567890"
Message = "test"

' I open a text file on COM1
Dim fso, f
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("COM14:", ForWriting, False, TriStateFalse)

' This is just an initialisation command for the Siemens
f.Write "AT+CMGF=1" & Chr(13)
WSCript.Sleep DelayMilliSecs

' This gives the phone number to the Siemens
f.Write "AT+CMGS=" &PhoneNum & Chr(13)
WSCript.Sleep DelayMilliSecs

' This sends the message
f.Write Message & Chr(13)
WSCript.Sleep DelayMilliSecs

' Siemens asks that you finish your message with Ctrl-Z
f.Write chr(26)
WSCript.Sleep DelayMilliSecs
