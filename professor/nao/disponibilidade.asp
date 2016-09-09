<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 1200
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a38")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Disponibilidade de Horário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open application("consql")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao2
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao2
set rsn=server.createobject ("ADODB.Recordset")
Set rsn.ActiveConnection = conexao

mes=request.form("mes")
mes=4
ano=request.form("ano")
ano=2005
udia=day(dateserial(ano,mes+1,1)-1)
sqln="select top 20 chapa from pfunc where codsindicato='03' and codsituacao in ('A','F') " & _
" and chapa in (select chapa from quem_nomeacoes) " & _
"order by chapa "
sqln="select top 100 chapa from pfunc where codsindicato='03' and codsituacao in ('A','F','Z') " & _
" and chapa in (select chapa from grades_rt where codevento in ('128','255','256','257','258','034') ) " & _
"order by chapa "
sqln="select top 10 chapa from pfunc where codsindicato='03' and codsituacao in ('A','F') " & _
" and chapa>'00000' " & _
"order by chapa "

rsn.Open sqln, ,adOpenStatic, adLockReadOnly
rsn.movefirst
do while not rsn.eof
chapa=numzero(request.form("chapa"),5)
chapa=rsn("chapa")
	
sqld="select f.nome, c.nome as funcao, s.descricao as setor, p.sexo from pfunc f, psecao s, pfuncao c, ppessoa p where f.codpessoa=p.codigo and f.codsecao=s.codigo and f.codfuncao=c.codigo and f.chapa='" & chapa & "'"
rsd.Open sqld, ,adOpenStatic, adLockReadOnly
nome=rsd("nome"):setor=rsd("setor"):funcao=rsd("funcao"):sexo=rsd("sexo")
rsd.close
linha=0
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campop" colspan=3><b>DISPONIBILIDADE DE HORÁRIO</td>
<td class=campo align="right"><%=rsn.absoluteposition%></td>
</tr>
<tr><td class=campo>Chapa</td>
	<td class=campo>Nome</td>
	<td class=campo>Setor</td>
	<td class=campo>Função</td></tr>
<tr><td class=campo><%=chapa%></td>
	<td class=campo><b><%=nome%></b></td>
	<td class=campo><%=Setor%></td>
	<td class=campo><%=funcao%></td></tr>
</table>
<table border="1" bordercolor="gray" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td height=25 class=titulo align="center">Dia da Semana</td>
	<td class=titulo align="center" colspan=6>Horários</td>
</tr>
<tr>
	<td height=25 class=titulo>Matutino</td>
	<td class=titulo align="center">7:30-8:20</td>
	<td class=titulo align="center">8:20-9:10</td>
	<td class=titulo align="center">9:20-10:10</td>
	<td class=titulo align="center">10:10-11:00</td>
	<td class=titulo align="center">11:10-12:00</td>
	<td class=titulo align="center">12:00-12:50</td>
</tr>
<%
for a=2 to 7
%>
<tr>
	<td height=25 class=campo><%=weekdayname(a)%></td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
</tr>
<%
next
%>
<tr>
	<td height=25 class=titulo>Vespertino</td>
	<td class=titulo align="center">13:00-13:50</td>
	<td class=titulo align="center">13:50-14:40</td>
	<td class=titulo align="center">14:50-15:40</td>
	<td class=titulo align="center">15:40-16:30</td>
	<td class=titulo align="center">16:40-17:30</td>
	<td class=titulo align="center">17:30-18:20</td>
</tr>
<%
for a=2 to 7
%>
<tr>
	<td height=25 class=campo><%=weekdayname(a)%></td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
</tr>
<%
next
%>
<tr>
	<td height=25 class=titulo>Noturno</td>
	<td class=titulo align="center">19:30-20:20</td>
	<td class=titulo align="center">20:20-21:10</td>
	<td class=titulo align="center">21:20-22:10</td>
	<td class=titulo align="center">22:10-23:00</td>
	<td class=campo colspan=2 rowspan=6>&nbsp;</td>
</tr>
<%
for a=2 to 6
%>
<tr>
	<td height=25 class=campo><%=weekdayname(a)%></td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
</tr>
<%
next
%>
</table>
<%
if sexo="M" then f1="o" else f1="a"
if sexo="M" then f2="" else f2="a"
%>
<p style="margin-top: 0; margin-bottom: 0"><font size=2>
<b>Prezad<%=f1%> Professor<%=f2%>, marque com um "X" a sua disponibilidade de horários para o próximo semestre.
</font></p>
<DIV style="page-break-after:always"></DIV>
<%
rsn.movenext
loop
rsn.close

set rs=nothing
set rs2=nothing
set rsd=nothing
set rsn=nothing
conexao.close
set conexao=nothing
conexao2.close
set conexao2=nothing
%>
</body>
</html>