<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Confirmação de presença</title>
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		if request.form("naovai")="on" then naovai=1 else naovai=0
		
		sql="UPDATE _conv SET "
		sql=sql & "dtconfirmacao='" & dtaccess(request.form("dtconfirmacao")) & "' "
		sql=sql & ", confirmado= '" & request.form("confirmado")& "' "
		sql=sql & ", naovai=" & naovai
		sql=sql & ", usuarioa='" & session("usuariomaster") & "' "
		sql=sql & " WHERE codigo=" & request.form("id_form") & " "
		response.write sql
		conexao.Execute sql, , adCmdText
		
		sql1="select tratamento, nome, descricao, local from _conv where codigo=" & request.form("id_form")
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		tratamento=rs("tratamento")
		nome=rs("nome")
		local=rs("descricao")
		rs.close
		
		Set Mailer = CreateObject("CDO.Message") 
		Mailer.From = "02379@unifieo.br" ' e-mail de quem esta enviando a mensagem 
		Mailer.To = "adriano.valentim@unifieo.br" ' e-mail de quem vai receber a mensagem 
		if  session("usuariomaster")="" then usuario="rh" else usuario=session("usuariomaster")
		Mailer.CC = usuario & "@unifieo.br, eleni@unifieo.br"
		'Mailer.BCC = "00259@unifieo.br" ' Com Cópia
		'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
		if naova1=1 then textoconf0="Ausência" else textoconf0="Presença"
		Mailer.Subject = "Confirmação de " & textoconf0 & " para a solenidade de entrega de titulo Honoris Causa"
		'Mailer.TextBody = "Você tem mensagem" 
		if naovai=1 then textoconf=" foi confirmada a <b>NÃO</b> presença " else textoconf=" foi confirmada a presença "
		Texto1="Informamos que em " & request.form("dtconfirmacao") & " por " & request.form("confirmado") & textoconf & " de "
		Texto2=tratamento & " " & nome & " (" & local & ")."
		
		Mailer.HtmlBody=Texto1 & Texto2 & "<br><br>" & _
		"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905</table>"
		'response.write "<br><br></table>" &Mailer.HtmlBody
		'if session("usuariomaster")="02379" then Mailer.Send
'==This section provides the configuration information for the remote SMTP server.
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.unifieo.br"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "02379@unifieo.br"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "123456"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
Mailer.Configuration.Fields.Update
'==End remote SMTP server configuration section==
		Mailer.Send
		Response.write "<font color=red> E"
		Set Mailer = Nothing 


	end if

	if request.form("bt_excluir")<>"" then
		'sql="DELETE id_form FROM uprofformacao_ WHERE id_form=" & session("id_alt_form")
		'conexao.Execute sql, , adCmdText
	end if

else 'request.form=""

	if request("codigo")=null then
		id_form=session("alt_conv")
	else
		id_form=request("codigo")
	end if
end if

if request.form="" then
session("alt_conv")=id_form

%>
<form method="POST" action="convites2.asp" name="form">
<input type="hidden" name="id_form" size="4" value="<%=session("alt_conv")%>" style="font-size: 8 pt">
<input type="hidden" name="dtconfirmacao" value="<%=now()%>">

<table border="0" cellpadding="3" cellspacing="0" width="370">
<tr>
	<td class=titulo>Data da confirmação</td></tr>
<tr>
	<td class=fundo><p class=realce><%=int(now())%></p></td></tr>
<tr>
	<td class=titulo>Confirmado por
	</td></tr>

<tr>
	<td class=fundo><input type="text" name="confirmado" size="40" value="">
		Não vai	<input type="checkbox" name="naovai" value="on">
</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="370">
<tr>
	<td class=titulo align="center">
	<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
	<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
<!--	<input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td> -->
	</tr>
</table>
</form>
<%
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
%>
<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.document.form.submit();self.close();</script>
<%
end if

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>