<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Seleção para imprimir Falta de Marcações e Justificativa</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
	dim conexao, conexao2, chapach, rs, rs2
	set conexao=server.createobject ("ADODB.Connection")
	conexao.Open application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	Set rs.ActiveConnection = conexao
	set rs2=server.createobject ("ADODB.Recordset")
	Set rs2.ActiveConnection = conexao
	set rs3=server.createobject ("ADODB.Recordset")
	Set rs3.ActiveConnection = conexao
	
if request.form="" then
chapa=request("chapa")
datai=request("datai")
dataf=request("dataf")
%>
<p class=titulo>Verificação de Falta de Marcações (Email)
<form method="POST" action="n3_email.asp">
<input type="hidden" name="chapa" value="<%=chapa%>">
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=campo colspan=3>
<%
sql1="select chapa, nome, secao from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
response.write rs("chapa") & "-<b>" & rs("nome")
rs.close
%>	
	</td>
<tr>
	<td class=titulo width=120>Data</td>
	<td class=titulo width=220>Marcações efetuadas</td>
	<td class=titulo></td>
</tr>
<%
vezes=0
sql2="select a.data, datepart(dw,a.data) as diasem, envio=max(c.dtenvio), tipo=max(c.tipo), vezes=count(c.dtenvio) " & _
"from _marcacoes_checagem a left join n3controle c on c.chapa=a.chapa and c.data=a.data " & _
"where a.chapa='" & chapa & "' and a.data between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' " & _
"group by a.chapa, a.data order by a.data" 
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof 
%>
<tr>
	<td class=campo align="center"><%=rs2("data")%> (<%=weekdayname(weekday(rs2("data")),1)%>)</td>
	<td class=campo>
<%
sql3="select batida from corporerm.dbo.abatfun where chapa='" & chapa & "' and data='" & dtaccess(rs2("data")) & "' order by batida"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
	batida=rs3("batida")
	hora=int(batida/60)
	minuto=batida-(hora*60)
	temp=numzero(hora,2) & ":" & numzero(minuto,2)
	response.write temp
	if rs3.absoluteposition<rs3.recordcount then response.write " - "
rs3.movenext
loop
else
	response.write "-"
end if
rs3.close
%>
	</td>
	<td class=campo>
	<input type="checkbox" name="emitir<%=vezes%>" value="ON" <%="checked"%> >
	<input type="hidden" name="id<%=vezes%>" value="<%=chapa%>">
	<input type="hidden" name="dt<%=vezes%>" value="<%=rs2("data")%>">
	</td>
</tr>	
<%
if rs2("vezes")>0 then
	if rs2("tipo")="E" then tipo="Email" else tipo="Formulário"
	if rs2("vezes")>1 then texto1="vezes" else texto1="vez"
%>
<tr>
	<td class=fundor colspan=3><font color=red><b>
	<%="Ultimo envio em " & rs2("envio") & " por " & tipo & " (" & rs2("vezes") & texto1 & ")"%>
	</b></font>
	</td>
</tr>
<%
if envios<rs2("vezes") then envios=rs2("vezes")
end if
%>

<%
vezes=vezes+1
rs2.movenext
loop
session("n3print")=vezes-1
end if 'rs2.recordcount>0
rs2.close
%>
<input type="hidden" name="envios" value="<%=envios%>">
<tr><td colspan=3 class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3">
</td></tr>
	</table>
<!-- final do quadro dos dias com marcações incompletas -->	
</form>
<hr>

<%
else 'request.form <>''

	chapa=request.form("chapa")
	vez=session("n3print")
	envios=request.form("envios")
	if envios="" or isnull(envios) then envios=0
	sql="delete from n3print where sessao='" & session.sessionid & "' ":conexao.execute sql
	for a=0 to vez
		pchapa=request.form("id" & a)
		pdata=request.form("dt" & a)
		emitir=request.form("emitir" & a)
		'response.write pchapa & " " & pdata & " " & emitir & "<br>"
		if emitir="ON" then
			sql="INSERT INTO n3print ( sessao, data, chapa ) SELECT '" & session.sessionid & "', '" & dtaccess(pdata) & "', '" & pchapa & "'"
			conexao.execute sql
		end if
	next
linha=0:pagina=0

sql1="select chapa, nome, codsecao, secao, sexo, codhorario, email from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs("sexo")="M" then s1="o" else s1="a"
if rs("sexo")="M" then s2="" else s2="a"
if rs("sexo")="M" then s3="o" else s3=""

sqlchefe="select distinct chapa from ( " & _
"select 'Chapa'=CHAPASUBST from corporerm.dbo.PSUBSTCHEFE s where CODSECAO='" & rs("codsecao") & "' and (DATAFIM is null or GETDATE() between DATAINICIO and DATAFIM) " & _
"union " & _
"select CHAPASUBST from corporerm.dbo.PSUBSTSUP s where CODSECAO='" & rs("codsecao") & "' and (DATAFIM is null or GETDATE() between DATAINICIO and DATAFIM) " & _
") c "
rs2.Open sqlchefe, ,adOpenStatic, adLockReadOnly
emailchefe=""
if rs2.recordcount>0 then
do while not rs2.eof
	emailchefe=emailchefe & rs2("chapa") & "@unifieo.br"
	if rs2.absoluteposition<rs2.recordcount then emailchefe=emailchefe & ","
rs2.movenext
loop
end if
rs2.close

cabecalho="<html><style type='text/css'>" & _
"<!--" & _
"td.titulo { font-size:8pt; font-family:tahoma; font-weight:bold; background-color:Silver; color:Black;} " & _
"td.titulop { font-size:10pt; font-family:tahoma; font-weight:bold; background-color:Silver; color:Black;} " & _
"td.campo { font-size:8pt; font-family:tahoma; font-weight:normal; background-color:White; font-size-adjust:inherit; font-stretch:inherit;} " & _
"td.campop { font-size:10pt; font-family:tahoma; font-weight:normal; background-color:White; font-size-adjust:inherit; font-stretch:inherit;} " & _
"td.campor { font-size:9px; font-family:tahoma; font-weight:normal; background-color:White; font-style:inherit; font-variant:normal; font-size-adjust:0; font-stretch:inherit;}" & _
"td.fundor { font-size:9px; font-family:tahoma; font-weight:normal; background-color:Silver; color:Black;} " & _
"p { font-size:10pt; font-family:tahoma; font-weight:normal;} " & _
"-->"&_
"</style><body>"

t1="<table border=0 cellpadding=1 cellspacing=0 style=""border-collapse: collapse"" width=720 > " & _
"<tr><td class=""campop"" align=""left"" valign=top height=62><img src=""http://rh.unifieo.br/images/logo_centro_universitario_unifieo_big.gif"" width=250 border=0></td></tr>" & _
"<tr><td class=""campop"" align=""right"">Osasco,&nbsp;" & day(now()) & " de " & monthname(month(now())) & " de " & year(now()) & "</td></tr>"

t1a="<tr><td class=""campop"" valign=top><table border='0' cellpadding='1' cellspacing='0' style='border-collapse: collapse' width='100%' >" & _
"<tr><td class=""campop"" valign=top> " & _
"<br>A" & s3 & " Sr" & s2 & ".<br>" & rs("nome") & "&nbsp;(" & rs("chapa") & "/" & rs("codhorario") & ")<br>Setor: " & rs("secao") & "<br><br> " & _
"Ref.: Falta de marcações em seu ponto eletrônico.<br></td>" & _
"<td class=campo valign=middle align=""center"">"

if envios=1 then t1b="<img src='http://rh.unifieo.br/images/n3_2solicitacao.png' border='0' alt='2a.Solicitação'>"
if envios>=2 then t1b="<img src='http://rh.unifieo.br/images/n3_avisofinal.png' border='0' alt='Aviso Final'>"

t1c="</td></tr></table></td></tr>" & _
"</tr><tr><td class=""campop"">&nbsp;</td></tr>"

t2="<tr><td class=""campop"" valign=top>" & _
"<p style=""margin-top:0;margin-bottom:0;text-align:justify""> " & _
"Após verificação nas marcações em seu ponto eletrônico no mês de <b>" & monthname(month(pdata)) & "</b>, constatamos que por algum motivo algumas marcações não foram registradas." & _
"</td></tr>"
t2=""

t3="<tr><td class=""campop"" valign=top><p style=""margin-top:0;margin-bottom:0;text-align:justify""> " & _
"Preencher o quadro 'Justificativa para Ausência de Marcação de Ponto' abaixo, e devolver no <b>prazo máximo de 48 horas</b> " & _
"ao Recursos Humanos, para regularização, ficando ciente de que as informações serão incluídas manualmente nas suas marcações de ponto e conferidas com outros controles eletrônicos disponíveis, como catraca eletrônica, controle de estacionamento, entre outros. " & _
"</td></tr> " & _
"<tr><td class=""campop"" height=5></td></tr>"
t3=""

t4="<tr><td class=""campop"" valign=top><p style=""margin-top:0;margin-bottom:0;text-align:justify""> " & _
"<b>Lembramos que a Portaria nº 269/2013-Reitoria de 28/10/2013, no seu item 6, regulamenta penalidades e limita o número de esquecimentos a 2 por mês." & _
"<br><font color=red>A falta de marcações impede a emissão do espelho do cartão de ponto, e os atrasos e faltas " & _
"existentes durante o fechamento da folha de pagamento serão considerados como aceitos e lançados." & _
"</td></tr>"

t5="<tr><td class=""campop"" valign=top> " & _
"<!-- quadro dos dias com marcações incompletas -->" & _
"<table border='1' bordercolor='#000000' cellpadding='4' cellspacing='0' style=""border-collapse: collapse"" width=600> " & _
"<tr><td class=titulop>Data das marcações</td> " & _
"<td class=titulop>Marcações efetuadas</td> " & _
"<td class=titulop>Data das marcações</td> " & _
"<td class=titulop>Marcações efetuadas</td></tr>"

sql2="select a.chapa, a.data, datepart(dw,a.data) as diasem " & _
"from _marcacoes_checagem a inner join n3print p on p.chapa=a.chapa and p.data=a.data " & _
"where a.chapa='" & chapa & "' " & _
"group by a.chapa, a.data order by a.data" 
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof 

if rs2.absoluteposition/2-int(rs2.absoluteposition/2)<>0 then t6=t6 & "<tr>"
t6=t6 & "<td class=""campop"" align=""center"">" & rs2("data") & " (" & weekdayname(weekday(rs2("data")),1) & ") </td>" & _
"<td class=""campop"">"

sql3="select batida from corporerm.dbo.abatfun where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs2("data")) & "' order by batida"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
	batida=rs3("batida")
	hora=int(batida/60)
	minuto=batida-(hora*60)
	temp=numzero(hora,2) & ":" & numzero(minuto,2)
	'response.write temp
	t6=t6 & temp
	'if rs3.absoluteposition<rs3.recordcount then response.write " - "
	if rs3.absoluteposition<rs3.recordcount then t6=t6 & " - "
rs3.movenext
loop
else
	t6=t6 & "-"
end if
rs3.close

if rs2.absoluteposition/2-int(rs2.absoluteposition/2)=0 then t6=t6 & "<tr>"

sqli="insert into n3controle (chapa, data, dtenvio, tipo) " & _
"select '" & chapa & "', '" & dtaccess(rs2("data")) & "', getdate(), 'E' "
'response.write "<br>" & sqli
conexao.execute sqli

rs2.movenext
loop

t6=t6 & "</table>" & _
"<!-- final do quadro dos dias com marcações incompletas -->" & _
"</td></tr>"

t7="<tr><td class=""campop"">&nbsp;</td></tr> " & _
"<tr><td class=""campop"">Atenciosamente,<br><br>Recursos Humanos<br><br></td></tr>" & _
"<tr><td class=""campop"">&nbsp;</td></tr>" & _
"</table>"

f1="<table border='0' cellpadding='1' cellspacing='0' style=""border-collapse: collapse"" width=650 height=450>" & _
"<tr><td class=""campop"" valign=top>" & _
"<!-- quadro formulario justificativa -->" & _
"<table style=""border-collapse: collapse"" border='1' bordercolor='#CCCCCC' cellpadding='2' width=600 cellspacing='0'>" & _
"<tr><td><img border='0' src=""http://rh.unifieo.br/images/logo_centro_universitario_unifieo_big.jpg"" width=225 height=50></td>" & _
"<td align=""center""><b><font size=2>Justificativa para Ausência de Marcação de Ponto</font></b></td>" & _
"</tr></table>"

f2="<table style=""border-collapse: collapse"" border='1' bordercolor='#CCCCCC' cellpadding='2' width='600' cellspacing='0'>" & _
"<tr><td class=""campop"" valign='top'><font size=1>Departamento:</font><br>" & rs("secao") & "</td>" & _
"<td class=""campop"" width='150' valign='top'><font size=1>Mês:</font><br>" & monthname(month(pdata)) & "</td>" & _
"<td class=""campop"" width='100' valign='top'><font size=1>Ano:</font><br><b>" & year(now()) & "</b></td></tr></table>"

f3="<table style=""border-collapse: collapse"" border='1' bordercolor='#CCCCCC' cellpadding='2' width=600 cellspacing='0'>" & _
"<tr><td class=""campop"" width='80' valign='top'><font size=1>Chapa:</font><br>" & chapa & "</td>" & _
"<td class=""campop"" valign='top'><font size=1>Nome do Funcionário:</font><br>" & rs("nome") & "</td></tr></table>"

f4="<table style=""border-collapse: collapse"" border='1' bordercolor='#CCCCCC' cellpadding='2' width=600 cellspacing='0'>" & _
"<tr><td width='100%' valign='top' colspan='13'><font size=1>Destina-se o presente controle a registrar informações do Empregado, " & _
"relativas aos dias e horário de trabalho face a justificativa assinalada. Fica ciente o empregado, e autoriza, que as informações serão " & _
"incluídas manualmente nas suas marcações de ponto e conferidas com outros controles eletrônicos disponíveis, como catraca eletrônica, controle de estacionamento, entre outros.</font></td></tr> " & _
"<tr><td class=fundor colspan=5 align=""center"" style=""border:1px solid #000000""><i><b>Informe apenas as marcações com problema</td> " & _
"<td class=fundor colspan=8 align=""center"" style=""border:1px solid #000000""><i><b>Assinale o motivo</td></tr>"

f5="<tr><td width='30' valign='middle' rowspan='2' align='center'><font size='1'>DIA</font></td> " & _
"<td width='60' valign='middle' rowspan='2' align='center'><font size='1'>Horário de Entrada</font></td> " & _
"<td            valign='top'    colspan='2' align='center'><font size='1'>Intervalo para refeição</font></td> " & _
"<td width='60' valign='middle' rowspan='2' align='center'><font size='1'>Horário de Saída</font></td> " & _
"<td            valign='middle' colspan='8' align='center'><font size='1'>Justificativa p/ Ausência</font></td> " & _
"</tr><tr> " & _
"<td width='60' valign='top' align='center'><font size='1'>Saída</font></td> " & _
"<td width='60' valign='top' align='center'><font size='1'>Retorno</font></td> " & _
"<td width='20' valign='top' align='center'><font size='1'>EM</font></td> " & _
"<td width='20' valign='top' align='center'><font size='1'>TE</font></td> " & _
"<td width='20' valign='top' align='center'><font size='1'>RD</font></td> " & _
"<td width='20' valign='top' align='center'><font size='1'>EX</font></td> " & _
"<td width='210'valign='top' align='center'><font size='1'>Outros</font></td></tr> "

f6=""
rs2.movefirst
do while not rs2.eof

f6=f6 & "<tr><td valign='top' height='25' align='center'>" & day(rs2("data")) & "</td> " & _
"<td valign='top' height='25'>&nbsp;</td><td valign='top' height='25'>&nbsp;</td> " & _
"<td valign='top' height='25'>&nbsp;</td><td valign='top' height='25'>&nbsp;</td> " & _
"<td valign='top' height='25'>&nbsp;</td><td valign='top' height='25'>&nbsp;</td> " & _
"<td valign='top' height='25'>&nbsp;</td><td valign='top' height='25'>&nbsp;</td> " & _
"<td valign='top' height='25'>&nbsp;</td></tr> " 

rs2.movenext
loop
rs2.close

f7="<tr><td valign='top' height='25'>&nbsp;</td><td valign='top' height='25'>&nbsp;</td> " & _
"<td valign='top' height='25'>&nbsp;</td><td valign='top' height='25'>&nbsp;</td> " & _
"<td valign='top' height='25'>&nbsp;</td><td valign='top' height='25'>&nbsp;</td> " & _
"<td valign='top' height='25'>&nbsp;</td><td valign='top' height='25'>&nbsp;</td> " & _
"<td valign='top' height='25'>&nbsp;</td><td valign='top' height='25'>&nbsp;</td></tr> " & _
"<tr><td valign='top' colspan='13'><font size='1'>Cód. Justificativas:<br> " & _
"<b>EM</b> - Esquecimento de marcação |  " & _
"<b>TE</b> - Trabalho Externo <b><i>(anexar relatório identificando)</i></b> | " & _
"<b>RD</b> - Relógio sem papel | " & _
"<b>EX</b> - Excluir marcação em excesso</font></td> " & _
"</tr></table> "

f8="<table style=""border-collapse: collapse"" border='1' bordercolor='#CCCCCC' cellpadding='2' width=600 cellspacing='0'> " & _
"<tr><td width=""100%""> " & _
"	<table style=""border-collapse: collapse""  border='0' cellpadding='0' width=""100%""> " & _
"	<tr> " & _
"		<td width=""30%"" class=""campor"" valign=bottom>&nbsp;<br>____________________<br>Data</td> " & _
"		<td width=""30%"" class=""campor"" valign=bottom>&nbsp;<br>_________________________<br>Assinatura do Funcionário</td> " & _
"		<td width=""40%"" class=""campor"" valign=bottom>De acordo:<br><br>&nbsp;&nbsp;&nbsp;__________________________________<br>&nbsp;&nbsp;&nbsp;Assinatura da Chefia</td> " & _
"	</tr></table> " & _
"</td></tr></table> "

f9="<table border='0' cellpadding='2' width='600' cellspacing='0'> " & _
"<tr><td width=""100%"" align=""right"" class=""campor"">Form.RH 11/2013</td> " & _
"</tr></table> " & _
"<!-- final do quadro formulario justificativa -->	" & _
"</td></tr> " & _
"<tr><td class=campo height='100%'>&nbsp;</tr> " & _
"</table>"

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

	Set Mailer = CreateObject("CDO.Message")
	Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = chapa & "@unifieo.br" ' e-mail de quem vai receber a mensagem 
	if rs("email")<>"" then Mailer.CC=rs("email")
	Mailer.BCC = emailchefe ' Com Cópia 
	Mailer.ReplyTo = "rh@unifieo.br"
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "Falta de Marcacao - Fazer justificativa urgente" 
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=cabecalho & t1 & t1a & t1b & t1c & t2 & t3 & t4 & t5 & t6 & t7 & _
	f1 & f2 & f3 & f4 & f5 & f6 & f7 & f8 & f9
'==This section provides the configuration information for the remote SMTP server.
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "02379@unifieo.br"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "eb541627"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
Mailer.Configuration.Fields.Update

'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.unifieo.br"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "02379@unifieo.br"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "*12345678"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
'Mailer.Configuration.Fields.Update

'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.hover.com"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "edson@benevides.com"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "queroserbenevides"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
'Mailer.Configuration.Fields.Update

'==End remote SMTP server configuration section==
	teste=0
	Mailer.Send 
	Set Mailer = Nothing 
	response.write "<p> Email enviado para " & chapa & "@unifieo.br" & " " & rs("email")

end if ' request.form	
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>