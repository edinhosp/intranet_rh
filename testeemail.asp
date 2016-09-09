<%@ Language=VBScript %>
<!-- #Include file="ADOVBS.INC" -->
<!-- #Include file="funcoesclear.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<img src="../images/logo_centro_universitario_unifieo_big.jpg" width="225" height="50" alt="">
</font>
<br><br><br>
<%
	dim conexao, conexao2
	dim rs, rs2
	set conexao=server.createobject ("ADODB.Connection")
	conexao.Open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	set rs.ActiveConnection = conexao
	set rs2=server.createobject ("ADODB.Recordset")
	set rs2.ActiveConnection = conexao
	set rst=server.createobject ("ADODB.Recordset")
	set rst.ActiveConnection = conexao

%>

<%

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

'******************* ROTINA PARA INFORMAR AOS PROFESSORES CARTA DE TETO ****************
hoje=formatdatetime(now,2)

cabecalho="<html><style type='text/css'>" & _
"<!--" & _
"td.titulo { font-size:8pt; font-family:tahoma; font-weight:bold; background-color:Silver; color:Black;} "& _
"td.campo { font-size:8pt; font-family:tahoma; font-weight:normal; background-color:White; font-size-adjust:inherit; font-stretch:inherit;} " & _
"p { font-size:10pt; font-family:tahoma; font-weight:normal;} " & _
"-->"&_
"</style><body>"
texto1="<table border='1' bordercolor='#a9a9a9' cellpadding='5' cellspacing='0' style='border-collapse: collapse' width=600'>" & _
"<tr><td class=campo><img src=""http://10.0.1.91/images/logo_centro_universitario_unifieo_big.gif"" width=230 border=0></td></tr>" & _
"<tr><td class=titulo>Esclarecimento</td></tr>" & _
"<tr><td class=campo>"
texto2="" 'nome
texto3="<p style='margin-bottom:0;margin-top:15'>Após um problema no último final de semana, nosso sistema repetidamente enviou a V.Sa. e-mails " & _
"sobre a declaração de contribuição ao INSS."
texto4="<p style='margin-bottom:0;margin-top:15'>Pedimos desculpas pelo transtorno e reiteramos que o incidente não voltará a acontecer. " & _
"<br><br>"

sql1="SELECT ct.chapa, f.NOME, p.SEXO, f.CODSITUACAO, p.EMAIL " & _
"FROM (rhcontroleteto AS ct INNER JOIN corporerm.dbo.pfunc AS f ON ct.chapa = f.CHAPA collate database_default) INNER JOIN corporerm.dbo.ppessoa AS p ON f.CODPESSOA = p.CODIGO " & _
"where cast((case mes when 13 then 12 else mes end) as char(2))+'/01/'+cast(ano as char) >= dateadd(m,-2,getdate()) and f.chapa not in ('00063') " & _
"GROUP BY ct.chapa, f.NOME, p.SEXO, f.CODSITUACAO, p.EMAIL HAVING f.CODSITUACAO<>'D' "
'sql1="select f.chapa, f.nome, p.sexo, p.email from pfunc f, ppessoa p where f.codpessoa=p.codigo and f.chapa='02379' "

rs.Open sql1, ,adOpenStatic, adLockReadOnly
if session("usuariomaster")="02379" then response.write " Teto:" & rs.recordcount
if rs.recordcount>0 then 
	do while not rs.eof
	if rs("sexo")="F" then pron="SRA. " else pron="SR. "
	texto2="<p style='margin-bottom:0;margin-top:15'><b>" & pron & rs("nome")
	if rs("email")<>"" then
		Set Mailer = CreateObject("CDO.Message") 
		Mailer.From = "02379@unifieo.br" ' e-mail de quem esta enviando a mensagem 
		Mailer.To = rs("email") '"movimentacao@unifieo.br" ' e-mail de quem vai receber a mensagem 
		Mailer.CC = rs("chapa") & "@unifieo.br"
		Mailer.BCC = "02379@unifieo.br,00259@unifieo.br" ' Com Cópia
		'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
		Mailer.Subject = "Esclarecimento sobre os inúmeros emails - " & monthname(month(now))
		'Mailer.TextBody = "Você tem mensagem" 
		Mailer.HtmlBody=cabecalho & texto1 & texto2 & texto3 & texto4 & "<br><br>" & _
		"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905</table>"
		'response.write "<br><br></table>" &Mailer.HtmlBody
'==This section provides the configuration information for the remote SMTP server.
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.unifieo.br"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "02379"
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

	rs.movenext
	loop
end if
rs.close


%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>