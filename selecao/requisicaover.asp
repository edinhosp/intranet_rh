<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a68")="N" or session("a68")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Requisição de Pessoal</title>
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

sqla="select * from rs_candidato "
sqlb="WHERE id_requisicao=" & request("codigo") & " "
sqlc="ORDER BY nome_candidato "

sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly
	
sql2="select * from rs_requisicao where id_requisicao=" & request("codigo")
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
id_requisicao=request("Codigo")	

%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
<% if session("a68")="T" then %>
<a href="requisicaover.asp?codigo=<%=id_requisicao%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover >
<img border="0" src="../images/write.gif" alt="Clique para atualizar">
<font size="1">!</font>
</a>
<% end if %>
REQUISIÇÃO DE PESSOAL</p>
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr><td class=grupo>Dados da Vaga/Requisição</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Descrição</td>
	<td class=titulor>&nbsp;Função</td>
</tr>
<%
sql3="select codigo, nome from corporerm.dbo.pfuncao where codigo='" & rs2("funcao") & "' "
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then funcao=rs3("nome") else funcao=""
rs3.close
vaga=rs2("descricao")
%>
<tr>
	<td class="campor"><b>&nbsp;<%=rs2("descricao")%></b></td>
	<td class="campor">&nbsp;<%=rs2("funcao")%> - <%=funcao%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Seção</td>
	<td class=titulor>&nbsp;Requisitante</td>
</tr>
<%
sql3="select codigo, descricao from corporerm.dbo.psecao where codigo='" & rs2("secao") & "' "
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then secao=rs3("descricao") else secao=""
rs3.close
sql3="SELECT nome from corporerm.dbo.pfunc where chapa='" & rs2("requisitante") & "' "
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then requisitante=rs3("nome") else requisitante=""
rs3.close
%>
<tr>
	<td class="campor">&nbsp;<%=rs2("secao")%> - <%=secao%></td>
	<td class="campor">&nbsp;<%=rs2("requisitante")%> - <%=requisitante%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Motivo</td>
	<td class=titulor>&nbsp;Funcionário substituído</td>
	<td class=titulor>&nbsp;Tipo</td>
	<td class=titulor>&nbsp;Cumprir Experiência</td>
</tr>
<%
select case rs2("motivo")
	case "02"
		motivo="Substituição"
	case "03"
		motivo="Vaga Nova"
	case "04"
		motivo="Aumento de quadro"
	case else
		motivo=""
end select
select case rs2("tipo")
	case "1"
		tipo="Normal"
	case "2"
		tipo="Estagiário"
	case else
		tipo=""
end select
sql3="SELECT nome from corporerm.dbo.pfunc where chapa='" & rs2("chapasubst") & "' "
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then chapasubst=rs3("nome") else chapasubst=""
rs3.close
%>
<tr>
	<td class="campor">&nbsp;<%=rs2("motivo")%> - <%=motivo%></td>
	<td class="campor">&nbsp;<%=rs2("chapasubst")%> - <%=chapasubst%></td>
	<td class="campor">&nbsp;<%=rs2("tipo")%> - <%=tipo%></td>
	<td class="campor">&nbsp;<%if rs2("exp_cumpre")=-1 then response.write "<font face='Wingdings'>ü</font>"%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor width=50>&nbsp;Salário</td>
	<td class=titulor width=50>&nbsp;Sal.Admissão</td>
	<td class=titulor>&nbsp;Horário</td>
</tr>
<%
sql3="SELECT codigo, descricao from corporerm.dbo.ahorario where codigo='" & rs2("horario") & "' "
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then horario=rs3("descricao") else horario=""
rs3.close
if rs2("salario")="" or isnull(rs2("salario")) then salario=0 else salario=cdbl(rs2("salario"))
if rs2("exp_cumpre")=-1 and rs2("tipo")=1 then fator=0.95 else fator=1
salario_exp=cdbl(salario*fator)
%>
<tr>
	<td class="campor" align="right">&nbsp;<%=formatnumber(salario,2)%></td>
	<td class="campor" align="right">&nbsp;<%=formatnumber(salario_exp,2)%></td>
	<td class="campor">&nbsp;<%=rs2("horario")%> - <%=horario%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr><td class=grupo>Requisitos da Vaga</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Escolaridade mínima</td>
	<td class=titulor>&nbsp;Idade</td>
	<td class=titulor>&nbsp;Sexo</td>
	<td class=titulor>&nbsp;Experiência</td>
	<td class=titulor>&nbsp;Deficiência</td>
	<td class=titulor>&nbsp;</td>
</tr>
<%
sql3="SELECT codcliente, descricao from corporerm.dbo.pcodinstrucao where codcliente='" & rs2("escolaridade") & "' "
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then escolaridade=rs3("descricao") else escolaridade=""
rs3.close
select case rs2("sexo")
	case "I"
		sexo="Indiferente"
	case "F"
		sexo="Feminino"
	case "M"
		sexo="Masculino"
	case else
		sexo=""
end select
select case rs2("deficiente")
	case "0"
		deficiente="Indiferente"
	case "1"
		deficiente="Não Deficiente"
	case "2"
		deficiente="Deficiente"
	case else
		deficiente=""
end select
%>
<tr>
	<td class="campor">&nbsp;<%=rs2("escolaridade")%> - <%=escolaridade%></td>
	<td class="campor">&nbsp;<%=rs2("idademin")%> min. <%=rs2("idademax")%> máx.</td>
	<td class="campor">&nbsp;<%=rs2("sexo")%> - <%=sexo%></td>
	<td class="campor">&nbsp;<%=rs2("experiencia")%> anos</td>
	<td class="campor">&nbsp;<%=rs2("deficiente")%> - <%=deficiente%></td>
	<td class="campor">&nbsp;<%=rs2("tp_def")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Cursos Exigidos</td>
	<td class=titulor>&nbsp;Outros Requisitos</td>
	<td class=titulor>&nbsp;Dt.Abertura</td>
	<td class=titulor>&nbsp;Dt.Encerr.</td>
	<td class=titulor>&nbsp;Qt.Vagas</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs2("cursos")%></td>
	<td class="campor">&nbsp;<%=rs2("outros")%></td>
	<td class="campor">&nbsp;<%=rs2("dt_abertura")%></td>
	<td class="campor">&nbsp;<%=rs2("dt_encerramento")%></td>
	<td class="campor">&nbsp;<%=rs2("qt_vagas")%></td>
</tr>
</table>

<!-- quadro inicio mudanca-->
<table border="1" bordercolor="#CCCCCC" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><th class=titulo colspan=9>Candidatos</th></tr>
<tr>
	<td class=titulor align="center" rowspan=2>Nome Candidato</td>
	<td class=titulor align="center" rowspan=2>Idade</td>
	<td class=titulor align="center" rowspan=2>Telefones</td>
	<td class=titulor align="center" rowspan=2>&nbsp;</td>
	<td class=titulor align="center" colspan=5>Agenda</td>
</tr>
<tr>
	<td class="campoa"r width=100>Processo</td>
	<td class="campoa"r width=40>Data</td>
	<td class="campoa"r width=40>Hora</td>
	<td class="campoa"r width=60>Observação</td>
	<td class="campoa"r width=20 align="center">i</td>
</tr>
<%
descricaovaga=rs2("descricao")
encerramento=rs2("dt_encerramento")
rs2.close
if rs.recordcount>0 then
'linhas=rs.recordcount
rs.movefirst
do while not rs.eof
sql3="SELECT * from rs_agenda where id_candidato=" & rs("id_candidato") & " order by processo_data"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
linhas=rs3.recordcount
if linhas=0 then linhas=1
%>
<tr>
	<td class="campor" rowspan=<%=linhas%> valign=middle> 
		<a class=r href="candidato_alteracao.asp?codigo=<%=rs("id_candidato")%>&vaga=<%=vaga%>" onclick="NewWindow(this.href,'AlteraCandidato','510','150','no','center');return false" onfocus="this.blur()">
		<%=rs("nome_candidato")%></a>
		<a class=r href="mailto:<%=rs("email")%>?subject=<%="Sobre a vaga " & descricaovaga%>" alt="Enviar email"><img src="../images/email_go.png" width="16" height="16" border="0" alt=""></a>
	</td>
	<td class="campor" rowspan=<%=linhas%> align="center"><%=rs("idade") %>    </td>
	<td class="campor" rowspan=<%=linhas%> align="left"><%=rs("telefone") %>   </td>
	<td class="campor" rowspan=<%=linhas%> align="center">
<%
sql2="select id_candidato, id_agenda, processo from rs_agenda where id_candidato=" & rs("id_candidato") & " and processo='7'"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
%>
	<a href="frm_contratacao.asp?codigo=<%=rs("id_candidato")%>" onclick="NewWindow(this.href,'frmContratacao','695','450','yes','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/note.gif" alt="imprimir contratação"></a>
<%
end if
rs2.close
%>
	</td>
<%
if rs3.recordcount>0 then
inicio=1
rs3.movefirst:do while not rs3.eof
sql2="select codigo, processo from rs_processo where codigo='" & rs3("processo") & "'"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then processo=rs2("processo") else processo=""
rs2.close
%>
	<td class="campor" width=100>
		<a class=r href="processo_alteracao.asp?codigo=<%=rs3("id_agenda")%>&candidato=<%=rs("nome_candidato")%>" onclick="NewWindow(this.href,'AlterarProcesso','510','150','no','center');return false" onfocus="this.blur()">
		<%=rs3("processo")%> - <%=processo%></a>
	</td>
	<td class="campor" width=40 align="center"><%=rs3("processo_data")%></td>
	<td class="campor" width=40 align="center"><%if rs3("processo_hora")<>"" then response.write formatdatetime(rs3("processo_hora"),4) else response.write "&nbsp;"%></td>
	<td class="campor" width=60><%=rs3("observacoes")%></td>
<% if inicio=1 then%>
	<td class="campor" width=20 rowspan=<%=linhas%>>
	<a href="processo_nova.asp?codigo=<%=rs("id_candidato")%>&candidato=<%=rs("nome_candidato")%>" onclick="NewWindow(this.href,'InclusaoProcesso','510','150','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/Appointment.gif" alt="inserir novo processo"></a>
	</td>
<%end if%>
</tr>
<%
inicio=0
rs3.movenext
loop
else
%>
	<td class="campor" width=240 colspan=4>&nbsp;</td>
	<td class="campor" width=20>
		<a href="processo_nova.asp?codigo=<%=rs("id_candidato")%>&candidato=<%=rs("nome_candidato")%>" onclick="NewWindow(this.href,'InclusaoProcesso','510','150','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/Appointment.gif" alt="inserir novo processo"></a>
	</td>
</tr>
<%
end if
rs3.close

rs.movenext
inicio=0
loop
else ' sem registros/planos
%>
<tr><td class="campor" colspan=8>&nbsp;</td></tr>
<%
end if
%>
</table>
<!-- quadro fim mudanca -->
<table><tr>
<td valign="top">
<% if session("a68")="T" then %>
<a href="candidato_nova.asp?codigo=<%=id_requisicao%>&vaga=<%=vaga%>" onclick="NewWindow(this.href,'InclusaoCandidato','510','150','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo candidato">
<font size="1">inserir novo candidato</font></a>
<% end if %>
</td>
</tr></table>
<hr>
<%
if encerramento<>"" then
%>
<a href="requisicaover.asp?codigo=<%=id_requisicao%>&enviaremail=S" >
<img src="../images/email_go.png" width="16" height="16" border="0" alt="Enviar email de agradecimento pela participação">
<font size="1">Enviar email de agradecimento</font></a>
<%
end if

if request("enviaremail")="S" then
adm="":dem=""
cabecalho="<html><style type='text/css'>" & _
"<!--" & _
"td.titulo { font-size:8pt; font-family:tahoma; font-weight:bold; background-color:Silver; color:Black;} "& _
"td.campo { font-size:8pt; font-family:tahoma; font-weight:normal; background-color:White; font-size-adjust:inherit; font-stretch:inherit;} " & _
"p { font-size:10pt; font-family:tahoma; font-weight:normal;} " & _
"-->"&_
"</style><body>"

dia=day(now())
mes=monthname(month(now()))
ano=year(now())

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

	Set Mailer = CreateObject("CDO.Message") 

sql1="select * from ( select r.id_requisicao, descricao, c.id_candidato, nome_candidato, email " & _
", aprovado=(select top 1 id_candidato from rs_agenda where id_candidato=c.id_candidato and processo=7) " & _
"from rs_requisicao r inner join rs_candidato c on c.id_requisicao=r.id_requisicao " & _
"where r.id_requisicao=" & id_requisicao & " and enviou=0 and email is not null ) z where aprovado is null "
rs3.Open sql1, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then 
	do while not rs3.eof

	intro="<table border='1' bordercolor='#a9a9a9' cellpadding='5' cellspacing='0' style='border-collapse: collapse' width=600'>" & _
	"<tr><td class=campo><img src='http://www.unifieo.br/images/logo-unifieo.png' border=0></td></tr>" & _
	"<tr><td class=campo>" & _
	"<p style='margin-bottom:0;margin-top:15'>Osasco, " & dia & " de " & mes & " de " & ano & "" & _
	"<br><br>" & _
	"<p style='margin-bottom:0;margin-top:15'>Prezado(a) " & rs3("nome_candidato") & "<br><br>" & _
	"<p style='margin-bottom:0;margin-top:15'>Agradecemos seu interesse em participar do nosso processo seletivo.<br><br>" & _
	"<p style='margin-bottom:0;margin-top:15'>Porém, após analisarmos seus dados frente às oportunidades oferecidas, concluímos não ser possível o seu aproveitamento nesta vaga.<br><br>" & _
	"<p style='margin-bottom:0;margin-top:15'>Desejamos muita sorte e comunicamos que seus dados serão armazenados em nosso arquivo, para uma próxima oportunidade.<br><br>" & _
	"<p style='margin-bottom:0;margin-top:15'>Mais uma vez, agradecemos o seu desempenho, sua colaboração, paciência e compreensão em todo processo seletivo.<br><br>" & _
	"<p style='margin-bottom:0;margin-top:15'>Atenciosamente,<br><br>"
	Mailer.From = "02379@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = rs3("email")         ' e-mail de quem vai receber a mensagem 
	Mailer.BCC = "02675@unifieo.br,02379@unifieo.br" ' Com Cópia 
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "A respeito da vaga de " & rs3("descricao") & "."
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=cabecalho & intro &	"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905 </td></tr></table></body></html>"
'==This section provides the configuration information for the remote SMTP server.
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.unifieo.br"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1'cdoBasic
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "02379@unifieo.br"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "123456"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
Mailer.Configuration.Fields.Update
'==End remote SMTP server configuration section==
	Mailer.Send

	sql1a="update rs_candidato set enviou=1 where id_candidato=" & rs3("id_candidato") & " and id_requisicao=" & rs3("id_requisicao") & ""
	conexao.execute sql1a
	rs3.movenext
	loop
end if
rs3.close
	Set Mailer = Nothing 

end if

%>

</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
set conexao=nothing
%>