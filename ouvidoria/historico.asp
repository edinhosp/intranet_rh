<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a37")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Histórico Escolar</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, sal_anterior(10)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
matricula=request("matricula")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

	
sql1="select e.matricula, e.idimagem, e.nome, e.estcivil, anoing, dtnasc, u.codcur, c.nome as curso, codtun, u.status, s.descricao " & _
"from corporerm.dbo.ealunos e, corporerm.dbo.ualucurso u, corporerm.dbo.ucursos c, corporerm.dbo.usitmat s " & _
"where e.matricula=u.mataluno and c.codcur=u.codcur and u.status=s.codsitmat and e.matricula='" & matricula & "' "
'"group by e.matricula, e.nome, anoing, e.estcivil, dtnasc, u.codcur, c.nome, codtun, u.status, s.descricao, pl.perletivo order by e.nome, pl.perletivo
rs.Open sql1, ,adOpenStatic, adLockReadOnly
imagem=rs("idimagem")
%>
<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=515>
<tr>
	<td class=titulor>Matricula</td>
	<td class=titulor>Nome</td>
	<td class=titulor>Est.Civil</td>
	<td class=titulor>Ano Ingr.</td>
	<td class=titulor>Dt.Nasc.</td>
	<td class=titulor>Curso</td>
	<td class=titulor>Status</td>
</tr>
<tr>
	<td class="campor"><%=rs("matricula")%></td>
	<td class="campor"><%=rs("nome")%></td>
	<td class="campor"><%=rs("estcivil")%></td>
	<td class="campor"><%=rs("anoing")%></td>
	<td class="campor"><%=rs("dtnasc")%></td>
	<td class="campor"><%=rs("curso")%></td>
	<td class="campor"><%=rs("descricao")%></td>
</tr>
</table>
	
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=515>
<tr>
	<td class=titulor>Periodo</td>
	<td class=titulor>Per</td>
	<td class=titulor>Status</td>
	<td class=titulor>&nbsp;</td>
</tr>
<%
rs.close
sql2="select perletivo, periodo, status, descricao from corporerm.dbo.umatricpl u, corporerm.dbo.usitmat s " & _
"where s.codsitmat=u.status and mataluno='" & matricula & "' "
rs.Open sql2, ,adOpenStatic, adLockReadOnly

if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class="campor"><%=rs("perletivo")%></td>
	<td class="campor"><%=rs("periodo")%></td>
	<td class="campor"><%=rs("descricao")%></td>
	<td class="campor">

<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulor>Disciplina</td>
	<td class=titulor>Média</td>
	<td class=titulor>Faltas</td>
	<td class=titulor>Freq.%</td>
	<td class=titulor>Status</td>
	<td class=titulor></td>
	<td class=titulor></td>
	<td class=titulor></td>
</tr>
<%		
sql3="select u.codmat, codtur, materia, a0,c0,f0,a1,c1,f1,a2,c2,f2,a3,c3,percfreq, u.cargahoraria, status, descricao " & _
"from corporerm.dbo.umatalun u, corporerm.dbo.umaterias m, corporerm.dbo.usitmat s " & _
"where u.codmat=m.codmat and s.codsitmat=u.status and u.mataluno='" & matricula & "' and u.perletivo='" & rs("perletivo") & "' "
rs2.Open sql3, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
c0="&nbsp;"
if not isnull(rs2("a0")) then a0=formatnumber(rs2("a0"),2) else a0=rs2("a0")
'if not isnull(rs2("c0")) then c0=formatnumber(rs2("c0"),2) else c0="&nbsp;"
if isnumeric(rs2("c0"))=true then c0=formatnumber(rs2("c0"),2)
if not isnull(rs2("c0")) and isnumeric(c0)=false then c0=rs2("c0")
if not isnull(rs2("percfreq")) then percfreq=formatnumber(rs2("percfreq"),2) else percfreq="&nbsp;"
if rs2("status")="08" or rs2("status")="09" or rs2("status")="10" then formato="<font color='red'>" else formato="<font color='black'>"
%>		
<tr>
	<td class="campor"><%=rs2("materia")%></td>
	<td class="campor" align="right"><%=a0%><%=c0%></td>
	<td class="campor" align="right"><%=rs2("f0")%>/<%=rs2("f1")%>-<%=rs2("f2")%></td>
	<td class="campor" align="right"><%=percfreq%></td>
	<td class="campor"><%=formato%><%=rs2("descricao")%></font></td>
	<td class="campor" align="right"><%=rs2("c1")%><%=rs2("a1")%></td>
	<td class="campor" align="right"><%=rs2("c2")%><%=rs2("a2")%></td>
	<td class="campor" align="right"><%=rs2("c3")%><%=rs2("a3")%></td>
</tr>
<%
codtur=rs2("codtur")
rs2.movenext
loop
end if
rs2.close
%>
</table><%=codtur%>
		</td>
	</tr>
<%
rs.movenext
loop
else
	response.write "<tr><td class=campo colspan=3>Sem lançamentos cadastrados</td></tr>"
end if
%>
</table>
<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>