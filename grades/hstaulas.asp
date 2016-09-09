<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Aulas atribuídas</title>
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

chapa=request("chapa")
inicio=request("inicio")
semestre=request("semestre")
perlet2=ano & "%" & semestre
	
sqla="SELECT g.*, h.descricao as descricao2 FROM g2ch g, g2defhor h " & _
"WHERE chapa1='" & chapa & "' and '" & dtaccess(inicio) & "' between inicio and termino " & _
"and g.turno=h.codtn and g.diasem=h.codds and g.pos=h.pos and deletada=0 AND H.TIPOCURSO=2 " & _
"order by diasem, turno, g.pos, coddoc " 
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=510>
	<th class=titulo colspan=7>Aulas atribuídas este semestre</th>
	<tr>
		<td class=titulor>Curso</td>
		<td class=titulor>Turma</td>
		<td class=titulor>Dia</td>
		<td class=titulor>Horário</td>
		<td class=titulor>J</td>
		<td class=titulor>Disciplina</td>
	</tr>
<%
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
%>
	<tr>
		<td class="campor"><%=rs("coddoc")%></td>
		<td class="campor"><%=rs("codtur")%></td>
		<td class="campor"><%=weekdayname(rs("diasem"),-1)%></td>
		<td class="campor" nowrap>&nbsp;<%=rs("descricao2")%></td>
		<td class=campo><%if rs("juntar")=true then response.write "<b>*</b>"%></td>
		<td class="campor"><%=rs("materia")%></td>
	</tr>
<%
rs.movenext:loop
%>
	<tr>
		<td class=grupo colspan=6><%=rs.recordcount%> aulas</td>
	</tr>
<%
else
	response.write "<tr><td class=campo colspan=3>Sem aulas atribuídas</td></tr>"
end if
%>
</table>
<br>
<%
rs.close
sqlb="select #=coddoc, Turma=codtur, Disciplina=materia+' ('+codmat collate database_Default+')' " & _
",'2014/2'=sum(case when perlet='2014/2' and termino='20150126' then ta else 0 end) " & _
",'2015/1'=sum(case when perlet='2015/1' and termino='20150730' then ta else 0 end) " & _
",'2015/2'=sum(case when perlet='2015/2' and termino='20160131' then ta else 0 end) " & _
",'2016/1'=sum(case when perlet='2016/1' and termino='20160731' then ta else 0 end) " & _
",'2016/2'=sum(case when perlet='2016/2' and termino='20170131' then ta else 0 end) " & _
"from g2ch where chapa1='" & chapa & "' and perlet in ('2014/2','2015/1','2015/2','2016/1','2016/2') " & _
"group by coddoc, codtur, codmat, materia "
rs.Open sqlb, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=510>
	<th class=titulo colspan=9>Comparativo</th>
	<tr>
		<td class=titulor>#</td>
		<td class=titulor>Turma</td>
		<td class=titulor>Disciplina</td>
		<td class=titulor>2014/2</td>
		<td class=titulor>2015/1</td>
		<td class=titulor>2015/2</td>
		<td class=titulor>2016/1</td>
		<td class=titulor>2016/2</td>
	</tr>
<%
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
%>
	<tr>
		<td class="campor"><%=rs("#")%></td>
		<td class="campor"><%=rs("turma")%></td>
		<td class="campor"><%=rs("disciplina")%></td>
		<td class="campor" align="center"><%=rs("2014/2")%></td>
		<td class="campor" align="center"><%=rs("2015/1")%></td>
		<td class="campor" align="center"><%=rs("2015/2")%></td>
		<td class="campor" align="center"><%=rs("2016/1")%></td>
		<td class="campor" align="center"><%=rs("2016/2")%></td>
	</tr>
<%
rs.movenext:loop
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