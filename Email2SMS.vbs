
' =========================================================
' Tittle: email 2 sms
' Version : v0.1
' Date: 04/08/2001
' Author: K Grey
' Comapny : None
' Comment: Send SMS via Email without Installing the SMTP Service
' Version History  :
' =========================================================

'this requires connection to the internet

Set objEmail = CreateObject("CDO.Message")

Const SMS = "@btpctext.co.uk"

Msgbox "Simple Demostration Of Email To SMS"

Number = InputBox("Enter Mobile Number","Email2SMS")
If Number = "" Then
	WScript.Quit
End If	

SMSMessage = InputBox("Enter Operator Message","Email2SMS")
If SMSMessage = "" Then
	WScript.Quit
End If	

objEmail.From = "my@email.co.uk"
objEmail.To = Number & SMS
objEmail.Subject = "BMS4 Test"
objEmail.Textbody = Now() & SMSMessage

objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

'Name or IP of Remote SMTP Server
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "email.domain.co.uk"

'Type of authentication, NONE, Basic (Base64 encoded), NTLM
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

'Your UserID on the SMTP server
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = "test user"

'Your password on the SMTP server
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "test passord"

'Server port (typically 25)
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 

'Use SSL for the connection (False or True)
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False


objEmail.Configuration.Fields.Update
objEmail.Send

Set MailObject = Nothing


