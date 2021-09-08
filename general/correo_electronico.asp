<%
If Request.ServerVariables("REQUEST_METHOD")="POST" Then
	'ENVÍO DEL FORMULARIO DE CONTACTO
	sch = "http://schemas.microsoft.com/cdo/configuration/"
	Set cdoConfig = CreateObject("CDO.Configuration")
	With cdoConfig.Fields
		.Item(sch & "sendusing") = 2
		'.Item(sch & "smtpserverpickupdirectory") = "C:\inetpub\mailroot\pickup" 
		.Item(sch & "smtpserver") = "mail.pisosypavimentos.com.ar"
		.Item(sch & "smtpserverport") = 587
		.Item(sch & "smtpconnectiontimeout") = 40
		.Item(sch & "smtpauthenticate") = 1
		.Item(sch & "sendusername") = "web@pisosypavimentos.com.ar"
		.Item(sch & "sendpassword") = "Nacho2018"
		.update
	End With

	Set MailObject = Server.CreateObject("CDO.Message")
	Set MailObject.Configuration = cdoConfig
	'MailObject.BodyFormat = 0
	'MailObject.mailformat = 0
	MailObject.From	= "web@pisosypavimentos.com.ar"
	MailObject.To	= "consultas@pisosypavimentos.com.ar"
	MailObject.Subject = "Consulta desde la Web"
	Cuerpo = "Nombre: " & Request.Form("nombre") & "<br>"
	Cuerpo = Cuerpo & "Email: " & Request.Form("email") & "<br>"
	Cuerpo = Cuerpo & "Tel: " & Request.Form("telefono") & "<br>"
	Cuerpo = Cuerpo & "Comentario: " & Request.Form("comentario") & "<br>"
	MailObject.HTMLBody = Cuerpo
	MailObject.Send
	Set MailObject = Nothing
	Set cdoConfig = Nothing
End If
%>