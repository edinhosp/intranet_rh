<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a91")="N" or session("a91")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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
ano=request("ano")
semestre=request("semestre")
perlet2=ano & "%" & semestre
	
sqla="SELECT g.*, h.descricao as descricao2 FROM grades_5ch g, grd_defhor h " & _
"WHERE chapa1='" & chapa & "' and perlet2 like '" & perlet2 & "' " & _
"and g.turno=h.codtn and g.diasem=h.codds and g.posicao=h.pos " & _
"order by diasem, turno, posicao, coddoc " 

rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=510>
	<th class=titulo colspan=7>Aulas atribuídas este semestre</th>
	<tr>
		<td class=titulor>Curso</td>
		<td class=titulor>Turma</td>
		<td class=titulor>Dia</td>
		<td class=titulor>Horário</td>
		<td class=titulor>Disciplina</td>
	</tr>
<%
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
%>
	<tr>
		<td class="campor"><%=rs("curso")%></td>
		<td class="campor"><%=rs("serie")&rs("turma")%></td>
		<td class="campor"><%=weekdayname(rs("diasem"),-1)%></td>
		<td class="campor" nowrap>&nbsp;<%=rs("descricao2")%></td>
		<td class="campor"><%=rs("materia")%></td>
	</tr>
<%
rs.movenext:loop
%>
	<tr>
		<td class=grupo colspan=5><%=rs.recordcount%> aulas</td>
	</tr>
<%
else
	response.write "<tr><td class=campo colspan=3>Sem aulas atribuídas</td></tr>"
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