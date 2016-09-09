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
<img src="images/logo_centro_universitario_unifieo_big.jpg" width="225" height="50" alt="">
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
	sql="SELECT Count(frase) total FROM frases"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	total=rs("total")
	rs.close
	randomize
	b=rnd(total)
	id=int(b*total)+1
	sql="select frase from frases where id=" & id
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	frase=rs("frase")
	end if
	rs.close
	sql="select frase from frases where id=" & total
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	frase2=rs("frase")
	end if
	rs.close
	if id=total then fundo="grupo" else fundo="titulo"
%>

<table border="1" bordercolor="'000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=300>
<tr>
	<td class=<%=fundo%>><%=frase%></td>
</tr>
</table>
<%
mesagora=month(now)
anoagora=year(now)
diaagora=day(now)

datasql=dateserial(anoagora,mesagora,diaagora)
incremento=0
datasql2=dateserial(anoagora,mesagora,diaagora+incremento)

sqla="SELECT a.*, u.nome from agenda a, usuarios u where u.usuario=a.usuarioc and a.tipo=0 and a.usuarioc='" & session("usuariomaster") & "' and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u where u.usuario=a.usuarioc and a.tipo=2 and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u where u.usuario=a.usuarioc and a.tipo=3 and a.usuarioc='" & session("usuariomaster") & "' and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u, agenda_3 a3 where a3.id_agenda=a.id_agenda and u.usuario=a.usuarioc and a.tipo=3 and a3.codigo='" & session("usuariomaster") & "' and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u, agenda_1 a1 where a1.id_agenda=a.id_agenda and u.usuario=a.usuarioc and a.tipo=1 and a1.codigo=u.grupo and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
'response.write sqla
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<script>
//Popup Window Script
//By JavaScript Kit (http://javascriptkit.com)
function openpopup(){
var popurl="agenda/popagenda.asp"
winpops=window.open(popurl,"popurl","width=430,height=200,scrollbars=no,center")
}
openpopup()
</script>
<%
end if
rs.close

sqlb="select top 1 datachecagem from agenda_check order by datachecagem desc "
rs.Open sqlb, ,adOpenStatic, adLockReadOnly
ultimadata0=rs("datachecagem"):ultimadata=ultimadata0+1
rs.close
hoje=formatdatetime(now,2)
if now()-ultimadata0>=7 then 'faz uma semana

	sql1="select f.chapa, f.nome, f.dtvencferias, dateadd(year,1,f.dtvencferias) as venc2, " & _
	"dateadd(year,1,f.dtvencferias)-60 as limite, f.inicprogferias1, (dateadd(year,1,f.dtvencferias)-60)-f.inicprogferias1 as retroativo " & _
	"from corporerm.dbo.pfunc f where f.codsituacao in ('A','F','Z') " & _
	"and dateadd(year,1,f.dtvencferias)-60 between '" & dtaccess(ultimadata) & "' and '" & dtaccess(hoje) & "' " & _
	"and (f.inicprogferias1 is null or dateadd(year,1,f.dtvencferias)-60-f.inicprogferias1<=-30) "
	rs2.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
	rs2.movefirst:do while not rs2.eof
		datap2=rs2("limite")
		if weekday(datap2)=7 then datap2=datap2-1
		if weekday(datap2)=1 then datap2=datap2-2
		datavenc2=rs2("limite")
		anotacao="Venc.2p.Férias <font color=red><b>" & datavenc2 & "</b></font> de " & rs2("nome")
		sqlf2="insert into agenda (data,compromisso,tipo,usuarioc,datac) select '" & dtaccess(datap2) & "'," & _
		"'" & anotacao & "',1,'99999','" & dtaccess(hoje) & "' " 
		'response.write sqlf2 & "<br>"
		conexao.Execute sqlf2, , adCmdText
	rs2.movenext:loop
	rs2.movefirst:do while not rs2.eof
		datap2=rs2("limite")
		if weekday(datap2)=7 then datap2=datap2-1
		if weekday(datap2)=1 then datap2=datap2-2
		datavenc2=rs2("limite")
		anotacao="Venc.2p.Férias <font color=red><b>" & datavenc2 & "</b></font> de " & rs2("nome")
		sqlr="select id_agenda from agenda where data='" & dtaccess(datap2) & "' and compromisso='" & anotacao & "' and usuarioc='99999' and datac='" & dtaccess(hoje) & "' and tipo=1 order by id_Agenda desc "
		rst.Open sqlr, ,adOpenStatic, adLockReadOnly
		idagenda=rst("id_agenda")
		rst.close
		sqli="insert into agenda_1 (id_agenda, codigo) select " & idagenda & ",'RH'; "
		'response.write sqli
		conexao.Execute sqli, , adCmdText
	rs2.movenext:loop
	end if
	rs2.close

	sqlf="insert into agenda_check (datachecagem) select '" & dtaccess(hoje) & "'; "
	response.write sqlf
	conexao.Execute sqlf, , adCmdText
else
	response.write "<font color=blue> A"
end if

'--------------------------------------- A V I S O S ------------------------------------
%>
<%
aviso=1
if aviso=1 then
%>
<bR>
<DIV align="center">
<table border="1" bordercolor="'000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=300>
<TR>
	<td class="campop">
	<br>
	<br>
	<br>
	O endereço para acessar o RH Online é <pre>http://rh.unifieo.br</pre>
	<br>
	<br>
	Este endereço (http://10.0.1.91) é para testes e pode ser interrompido a qualquer momento
	e está sujeito a falhas e interrupções.
	<br>
	</td>
</tr>
</table>
</div>
<%
end if
%>	


<%
'******************* ROTINA PARA INFORMAR ADMITIDOS E DEMITIDOS PARA CPD ****************
adm="":dem=""
cabecalho="<html><style type='text/css'>" & _
"<!--" & _
"td.titulo { font-size:8pt; font-family:tahoma; font-weight:bold; background-color:Silver; color:Black;} "& _
"td.campo { font-size:8pt; font-family:tahoma; font-weight:normal; background-color:White; font-size-adjust:inherit; font-stretch:inherit;} " & _
"p { font-size:10pt; font-family:tahoma; font-weight:normal;} " & _
"-->"&_
"</style><body>"
intro="<table border='1' bordercolor='#a9a9a9' cellpadding='5' cellspacing='0' style='border-collapse: collapse' width=600'>" & _
"<tr><td class=campo><img src='http://www.unifieo.br/images/logo-unifieo.png' border=0></td></tr>" & _
"<tr><td class=titulo>Movimentação de Funcionários</td></tr>" & _
"<tr><td class=campo>" & _
"<p style='margin-bottom:0;margin-top:15'>Estamos enviando as movimentações de entrada e saída de funcionários que ocorreram desde o último email." & _
"<p style='margin-bottom:0;margin-top:15'>Os desligados devem ser desativados nos serviços de rede e email, e os admitidos cadastrados.<br><br>"
introb="<table border='1' bordercolor='#a9a9a9' cellpadding='5' cellspacing='0' style='border-collapse: collapse' width=600'>" & _
"<tr><td class=campo><img src='http://www.unifieo.br/images/logo-unifieo.png' border=0></td></tr>" & _
"<tr><td class=titulo>Movimentação de Funcionários</td></tr>" & _
"<tr><td class=campo>" & _
"<p style='margin-bottom:0;margin-top:15'>Estamos enviando as movimentações de saída de funcionários que ocorreram desde o último email.<br><br>"
'"<p style='margin-bottom:0;margin-top:15'>Por gentileza, checar se há débitos existentes para estes funcionários."<br><br>"
introc="<table border='1' bordercolor='#a9a9a9' cellpadding='5' cellspacing='0' style='border-collapse: collapse' width=600'>" & _
"<tr><td class=campo><img src='http://www.unifieo.br/images/logo-unifieo.png' border=0></td></tr>" & _
"<tr><td class=titulo>Movimentação de Funcionários</td></tr>" & _
"<tr><td class=campo>" & _
"<p style='margin-bottom:0;margin-top:15'>Estamos enviando as movimentações de saída de professores que ocorreram desde o último email.<br><br>"
introc2="<table border='1' bordercolor='#a9a9a9' cellpadding='5' cellspacing='0' style='border-collapse: collapse' width=600'>" & _
"<tr><td class=campo><img src='http://www.unifieo.br/images/logo-unifieo.png' border=0></td></tr>" & _
"<tr><td class=titulo>Movimentação de Funcionários</td></tr>" & _
"<tr><td class=campo>" & _
"<p style='margin-bottom:0;margin-top:15'>Estamos enviando as movimentações de entrada de professores que ocorreram desde o último email.<br><br>"
introd="<table border='1' bordercolor='#a9a9a9' cellpadding='5' cellspacing='0' style='border-collapse: collapse' width=600'>" & _
"<tr><td class=campo><img src='http://www.unifieo.br/images/logo-unifieo.png' border=0></td></tr>" & _
"<tr><td class=titulo>Movimentação de Funcionários</td></tr>" & _
"<tr><td class=campo>" & _
"<p style='margin-bottom:0;margin-top:15'>Estamos enviando as movimentações de saída de funcionários que devem estar com bolsas de estudos ativas e que devem ser canceladas.<br><br>"

teveprofessor=0
entrouprofessor=0
tevebolsa=0
teste=0
SQL1="SELECT f.CHAPA, f.NOME, f.DATAADMISSAO AS ADMISSAO, s.DESCRICAO AS SECAO, p.CARTIDENTIDADE, " & _
"tipo=case when f.codtipo='T' then 'ESTAGIÁRIO' when f.codsindicato='03' then 'PROFESSOR' else 'ADMINISTRATIVO' END " & _
"FROM (infracpd_adm i RIGHT JOIN corporerm.dbo.PFUNC f ON i.CHAPA=f.CHAPA COLLATE DATABASE_DEFAULT) INNER JOIN corporerm.dbo.PSECAO s ON f.CODSECAO=s.CODIGO " & _
"inner join corporerm.dbo.PPESSOA p on p.CODIGO=f.codpessoa " & _
"WHERE i.CHAPA Is Null AND (f.CHAPA<'10000' Or f.CHAPA>'90000') and f.dataadmissao<getdate() and (f.tipodemissao<>'5' or f.TIPODEMISSAO is null)  "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if session("usuariomaster")="02379" then response.write " Adm:" & rs.recordcount
if rs.recordcount>0 then 
	adm="<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='100%'>" & _
	"<tr><td class=titulo colspan=5>Funcionários Admitidos</td></tr>" & _
	"<tr><td class=titulo>Registro</td><td class=titulo>Nome</td><td class=titulo>Admissão</td><td class=titulo>Setor/Seção</td><td class=titulo>Tipo</td></tr>"
	admp=adm
	do while not rs.eof
	if rs("tipo")="PROFESSOR" then
		admp=admp & "<tr><td class=""campo"">" & rs("chapa") & "</td><td class=""campo"">" & rs("nome") & "</td><td class=""campo"">" & rs("admissao") & "</td><td class=""campo"">" & rs("secao") & "</td><td class=""campo"">" & rs("tipo") & "</td></tr>"
		entrouprofessor=1
	end if
	adm=adm & "<tr><td class=""campo"">" & rs("chapa") & "</td>"
	adm=adm & "<td class=""campo"">" & rs("nome") & "<br>&nbsp;(" & rs("cartidentidade") & ")</td>"
	adm=adm & "<td class=""campo"">" & rs("admissao") & "</td>"
	adm=adm & "<td class=""campo"">" & rs("secao") & "</td>"
	adm=adm & "<td class=""campo"">" & rs("tipo") & "</td></tr>"
	sql1a="insert into infracpd_adm (chapa) select '" & rs("chapa") & "' "
	conexao.execute sql1a
	rs.movenext
	loop
	adm=adm & "</table>"
	admp=admp & "</table>"
end if
rs.close

if month(now())=12 or month(now())=6 then delay=1 else delay=0
sql2="SELECT f.CHAPA, f.NOME, f.DATADEMISSAO AS DEMISSAO, s.DESCRICAO AS SECAO, p.CARTIDENTIDADE, " & _
"tipo=case when f.codtipo='T' then 'ESTAGIÁRIO' when f.codsindicato='03' and f.chapa<'90000' then 'PROFESSOR' when f.codsindicato='03' and f.chapa>='98000' then 'CONVIDADO' else 'ADMINISTRATIVO' END, tipodemissao " & _
"FROM infracpd_dem I RIGHT JOIN (corporerm.dbo.PFUNC f INNER JOIN corporerm.dbo.PSECAO s ON f.CODSECAO=s.CODIGO) ON I.CHAPA=f.CHAPA collate database_default " & _
"inner join corporerm.dbo.PPESSOA p on p.CODIGO=f.codpessoa " & _
"WHERE (f.CHAPA<'10000' Or f.CHAPA>'90000') AND f.DATADEMISSAO Is Not Null AND I.CHAPA Is Null and f.datademissao<getdate()+1+" & delay & " and (f.tipodemissao<>'5' or f.TIPODEMISSAO is null) "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
if session("usuariomaster")="02379" then response.write " Dem:" & rs.recordcount
if rs.recordcount>0 then 
	dem="<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='100%'>" & _
	"<tr><td class=titulo colspan=5>Funcionários Desligados</td></tr>" & _
	"<tr><td class=titulo>Registro</td><td class=titulo>Nome</td><td class=titulo>Saída</td><td class=titulo>Setor/Seção</td><td class=titulo>Tipo</td></tr>"
	demp=dem
	demb=dem
	do while not rs.eof
	if rs("tipo")="PROFESSOR" then
		demp=demp & "<tr><td class=""campo"">" & rs("chapa") & "</td><td class=""campo"">" & rs("nome") & "</td><td class=""campo"">" & rs("demissao") & "</td><td class=""campo"">" & rs("secao") & "</td><td class=""campo"">" & rs("tipo") & "</td></tr>"
		teveprofessor=1
	end if
	if rs("tipodemissao")="4" then
	sqlbolsa="select b.chapa, nome_bolsista, matricula, b.curso, parentesco from bolsistas b, bolsistas_lanc l where b.id_bolsa=l.id_bolsa and validade>'" & dtaccess(rs("demissao")) & "' and chapa='" & rs("chapa") & "'"
	rs2.Open sqlbolsa, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
		do while not rs2.eof
		demb=demb & "<tr><td class=""campo"">" & rs("chapa") & "</td><td class=""campo"">" & rs2("nome_bolsista") & "</td><td class=""campo"">" & rs2("matricula") & "</td><td class=""campo"">" & rs2("curso") & "</td><td class=""campo"">" & rs2("parentesco") & "</td></tr>"
		rs2.movenext
		loop
		tevebolsa=1
	end if
	rs2.close
	end if 'if tipo 4
	dem=dem & "<tr><td class=""campo"">" & rs("chapa") & "</td>"
	dem=dem & "<td class=""campo"">" & rs("nome") & "</td>"
	dem=dem & "<td class=""campo"">" & rs("demissao") & "</td>"
	dem=dem & "<td class=""campo"">" & rs("secao") & "</td>"
	dem=dem & "<td class=""campo"">" & rs("tipo") & "</td></tr>"
	sql2a="insert into infracpd_dem (chapa) select '" & rs("chapa") & "' "
	conexao.execute sql2a
	rs.movenext
	loop
	dem=dem & "</table>"
	demp=demp & "</table>"
	demb=demb & "</table>"
end if
rs.close

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

if len(adm)>0 or len(dem)>0 then 'tem movimentação
	Set Mailer = CreateObject("CDO.Message") 
	Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = "movimentacao@unifieo.br" ' e-mail de quem vai receber a mensagem 
	Mailer.BCC = "rh@unifieo.br,02555@unifieo.br" ' Com Cópia 
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "RH - Mensagem com Movimentação" 
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=cabecalho & intro & adm & "<br><br>" & dem & _
	"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905"
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

'==End remote SMTP server configuration section==
	if teste=0 then Mailer.Send 
	Set Mailer = Nothing 
end if 'tem movimentação

if len(dem)>0 then 'tem movimentação
	Set Mailer = CreateObject("CDO.Message") 
	Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = "02283@unifieo.br" ' e-mail de quem vai receber a mensagem 
	'Mailer.CC = "02379@unifieo.br" ' Com Cópia 
	Mailer.BCC = "rh@unifieo.br" ' Com Cópia 
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "RH - Saída de Funcionários - Checagem de Débitos" 
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=cabecalho & introb & "<br><br>" & dem & _
	"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905"
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

'==End remote SMTP server configuration section==
	if teste=0 then Mailer.Send 
	Set Mailer = Nothing 
end if 'tem movimentação

if len(demp)>0 and teveprofessor=1 then 'tem movimentação
	Set Mailer = CreateObject("CDO.Message") 
	Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = "planejamento@unifieo.br" ' e-mail de quem vai receber a mensagem 
	Mailer.CC = "03210@unifieo.br" ' Com Cópia 
	Mailer.BCC = "rh@unifieo.br" ' Com Cópia 
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "RH - Saída de Professores" 
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=cabecalho & introc & "<br><br>" & demp & _
	"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905"
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

'==End remote SMTP server configuration section==
	if teste=0 then Mailer.Send 
	Set Mailer = Nothing 
end if 'tem movimentação

if len(admp)>0 and entrouprofessor=1 then 'tem movimentação
	Set Mailer = CreateObject("CDO.Message") 
	Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = "planejamento@unifieo.br" ' e-mail de quem vai receber a mensagem 
	Mailer.CC = "03210@unifieo.br" ' Com Cópia 
	Mailer.BCC = "rh@unifieo.br" ' Com Cópia 
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "RH - Entrada de Professores"
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=cabecalho & introc2 & "<br><br>" & admp & _
	"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905"
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

'==End remote SMTP server configuration section==
	if teste=0 then Mailer.Send 
	Set Mailer = Nothing 
end if 'tem movimentação


if len(demb)>0 and tevebolsa=1 then 'tem movimentação
	Set Mailer = CreateObject("CDO.Message") 
	Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	if teste=0 then Mailer.To = "00942@unifieo.br" ' e-mail de quem vai receber a mensagem 
	if teste=1 then Mailer.To = "02379@unifieo.br" 'teste
	if teste=0 then Mailer.CC = "slinos@unifieo.br" 'com copia
	Mailer.BCC = "rh@unifieo.br" ' Com Cópia 
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "RH - Saída de Funcionários com bolsas de estudos" 
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=cabecalho & introd & "<br><br>" & demb & _
	"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905"
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

'==End remote SMTP server configuration section==
	if teste=0 then Mailer.Send 
	Set Mailer = Nothing 
end if 'tem movimentação


'******************* ROTINA PARA RH - EMAIL ACHAPA ****************
teste=1
response.cookies("intranet_rh")("email_nom")="N"
'response.write ">>>>>" & request.cookies("intranet_rh")("email_nom")
if teste=0 and request.cookies("intranet_rh")("email_nom")<>"S" and session("usuariomaster")="02379" then
	'nao gerou
	response.cookies("intranet_rh")("email_nom")="S"
	response.cookies("intranet_rh").expires=dateadd("d",1,now)

sql1="select chapa, nome, email, sexo from qry_funcionarios f where chapa collate database_default in ('02258','02241','00657','00710','00745') and f.codsituacao<>'D' "
sql1="select chapa, nome, email, sexo from qry_funcionarios f where chapa collate database_default in (select chapa1 from achapa) and f.codsituacao<>'D' order by chapa "
sql1="select f.chapa, f.nome, f.email, f.sexo, a.observacao from qry_funcionarios f inner join achapa a on a.chapa1=f.chapa collate database_default where f.codsituacao<>'D' order by chapa "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then 
	do while not rs.eof

cabecalho="<html><style type='text/css'>" & _
"<!--" & _
"td.titulo { font-size:8pt; font-family:tahoma; font-weight:bold; background-color:Silver; color:Black;} "& _
"td.campo { font-size:8pt; font-family:tahoma; font-weight:normal; background-color:White; font-size-adjust:inherit; font-stretch:inherit;} " & _
"p { font-size:10pt; font-family:tahoma; font-weight:normal;} " & _
"-->"&_
"</style><body>"
texto1="<table border='1' bordercolor='#a9a9a9' cellpadding='5' cellspacing='0' style='border-collapse: collapse' width=600'>" & _
"<tr><td class=campo><img src='http://www.unifieo.br/templates/unifieonew/images/LGO.png' border=0></td></tr>" & _
"<tr><td class=titulo>Adendo ao Contrato de Trabalho</td></tr>" & _
"<tr><td class=campo>"
texto2="<p style='margin-bottom:0;margin-top:15'>" & rs("nome") & "<br>"
if rs("sexo")="F" then saudacao="Prezada Professora" else saudacao="Prezado Professor"
texto3="<p style='margin-bottom:0;margin-top:15'>" & saudacao & "<br>Comparecer ao Recursos Humanos para assinar adendo ao contrato de trabalho ref " & rs("observacao") & ", pendente de assinatura."
texto4="<p style='margin-bottom:0;margin-top:15'><br>Atenciosamente<br>"
intronom=texto1 & texto2 & texto3 & texto4
nom="<p style='margin-bottom:0;margin-top:15'><br>"
	Set Mailer = CreateObject("CDO.Message") 
	Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = rs("chapa")& "@unifieo.br" ' e-mail de quem vai receber a mensagem 
	if rs("email")<>"" then Mailer.CC = rs("email") else Mailer.CC="" 'com copia
	'Mailer.BCC = "rh@unifieo.br" ' Com Cópia 
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "Assinatura de Adendo" 
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=cabecalho & intronom & "<br><br>" & _
	"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905"
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

'==End remote SMTP server configuration section==
testee=0
	if testee=0 then Mailer.Send 
	Set Mailer = Nothing 

	rs.movenext
	loop

rs.close

else
	'ja gerou
	'response.cookies("intranet_rh")("gerouhoje")="N"
end if

end if


%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>