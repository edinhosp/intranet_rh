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
<title>Disciplinas e Professores</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, t(4)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
codmat=request("codmat")
sql1="select materia from grades_materias where codmat='" & codmat & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then materia=rs("materia") else materia=""
rs.close

sqli="select top 4 perlet3 from grades_5ch group by perlet3 order by perlet3 desc"
a=1
rs.Open sqli , ,adOpenStatic , adLockReadOnly
do while not rs.eof
	t(a)=rs("perlet3")
	a=a+1
rs.movenext
loop
rs.close
	
sqla="SELECT g.chapa1, d.NOME, d.TELEFONE1, g.materia, min(t.[" & t(4) & "]) as t1, min(t.[" & t(3) & "]) as t2, min(t.[" & t(2) & "]) as t3, min(t.[" & t(1) & "]) as t4 " & _
"FROM (grades_5ch g INNER JOIN totalizador_5ch t ON g.chapa1 = t.chapa1) INNER JOIN dc_professor d ON g.chapa1 = d.CHAPA collate database_default " & _
"WHERE g.codmat='" & codmat & "' AND d.CODSITUACAO In ('A','F','Z') " & _
"GROUP BY g.chapa1, d.NOME, d.TELEFONE1, g.materia " & _
"HAVING g.materia='" & materia & "' "	
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=510>
<th class=titulo colspan=7>Professores com " <%=materia%> "</th>
<tr>
	<td class=titulor rowspan=2>Chapa</td>
	<td class=titulor rowspan=2>Nome</td>
	<td class=titulor rowspan=2>Telefone</td>
	<td class=titulor colspan=3 align="center">Nº Aulas</td>
</tr>
<tr>
	<td class=titulor align="center"><%=t(3)%></td>
	<td class=titulor align="center"><%=t(2)%></td>
	<td class=titulor align="center"><%=t(1)%></td>
<%
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
%>
<tr>
	<td class="campor"><%=rs("chapa1")%></td>
	<td class="campor"><%=rs("nome")%></td>
	<td class="campor"><%=rs("telefone1")%></td>
	<td class="campor" align="center"><%=rs("t3")%></td>
	<td class="campor" align="center"><%=rs("t2")%></td>
	<td class="campor" align="center"><%=rs("t1")%></td>
</tr>
<%
rs.movenext:loop
%>
<tr>
	<td class=grupo colspan=6><%=rs.recordcount%> professores</td>
</tr>
<%
else
	response.write "<tr><td class=campo colspan=6>Sem aulas atribuídas</td></tr>"
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