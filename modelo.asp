<!-- #config timefmt="%m/%d/%y" -->
<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
'	Response.buffer=true
'	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.redirect "intranet.asp"
if session("a1")="N" or session("a1")="" then response.redirect "intranet.asp"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Grade Horária</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->

<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
'sqla="SELECT dc_carga.CURSO FROM dc_carga GROUP BY dc_carga.CURSO;"
'rs.Open sql1, ,adOpenStatic, adLockReadOnly


for each item in Request.ServerVariables
	Response.write item & " = " & request.servervariables(item) & "<br>"

next
response.write "<br>"

Response.Write Request.ServerVariables("SERVER_SOFTWARE") 


Set Mailer = CreateObject("CDO.Message") 
Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
Mailer.To = "02552@unifieo.br" ' e-mail de quem vai receber a mensagem 
Mailer.CC = "02379@unifieo.br" ' Com Cópia 
'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
Mailer.Subject = "Email Intranet - Mensagem Automática" 
'Mailer.TextBody = "Você tem mensagem" 
Mailer.HtmlBody="<b>Não responder.<br></b>Este email foi enviado pela Intranet."
Mailer.Send 
Set Mailer = Nothing 



%>


<%
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>

<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>