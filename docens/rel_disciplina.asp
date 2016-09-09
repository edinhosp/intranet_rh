<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a19")="N" or session("a19")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pesquisa de Disciplina</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, rs, chapach
dim filtro,filtro2,selecao,chave,palavra

set conexao=server.createobject ("ADODB.Connection")
conexao.open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form="" then
%>
<!-- selecoes -->
<form method="POST" action="rel_disciplina.asp" name="form">
  <p><font color="#0000ff"><b>Seleções para o relatório &quot;Pesquisa de
  Disciplinas&quot;</b></font></p>
  <p>Palavras chaves nas disciplinas a pesquisar:<br>
  palavra 1 <input type="text" name="T1" size="20" class="form_box"><br>
  palavra 2 <input type="text" name="T2" size="20" class="form_box"><br>
  palavra 3 <input type="text" name="T3" size="20" class="form_box"><br>
  palavra 4 <input type="text" name="T4" size="20" class="form_box"><br>
  palavra 5 <input type="text" name="T5" size="20" class="form_box"><br>
  palavra 6 <input type="text" name="T6" size="20" class="form_box"></p>
  <p><input type="submit" class=button value="Visualizar Relatório" name="B1">
  &nbsp;</p>
</form>

<p><font color="#FF0000">Configure a página do seu navegador (Internet
Explorer, Netscape, Mozilla, etc) no sentido RETRATO.</font></p>
<!-- modelo relatorio inicio -->
<!-- modelo relatorio inicio -->
<p>&nbsp;</p>
<%
end if 'if do request.form

if request.form<>"" then
%>
<%
filtro="":filtro2="":selecao="":dim t(4)
chave=0
for a=1 to 6
	palavra=request.form("T" & a)
	if palavra<>"" then
		if chave=0 then filtro="HAVING ":selecao="Seleção: disciplinas relacionadas com "
		if chave=1 then filtro=filtro & " or ":selecao=selecao & ", "
		filtro=filtro & "g.materia Like '%" & palavra & "%' " : selecao=selecao & palavra
		if chave=0 then chave=1
	end if
next
'if session("usuariomaster")="02379" then response.write "<br>" & filtro
sqli="select top 4 perlet from g2ch where perlet not like '%0' group by perlet order by perlet desc"
a=1
rs.Open sqli , ,adOpenStatic , adLockReadOnly
do while not rs.eof
	t(a)=rs("perlet")
	a=a+1
rs.movenext
loop
rs.close
'if session("usuariomaster")="02379" then response.write "<br>" & t(1)

sqla="SELECT g.materia, g.coddoc curso, d.chapa, d.nome, d.dataadmissao, d.codsituacao, d.instrucaomec, min(t.[" & t(4) & "]) as t1, min(t.[" & t(3) & "]) as t2, min(t.[" & t(2) & "]) as t3, min(t.[" & t(1) & "]) as t4 " & _
"FROM (g2ch AS g INNER JOIN dc_professor d ON g.chapa1 collate database_default= d.CHAPA) LEFT JOIN [totalizador_chor] AS t ON d.CHAPA = t.chapa1 collate database_default " & _
"WHERE g.deletada=0 and g.ativo=1 and d.codsituacao in ('A','F','Z') "
sqlb=""
sqlc="GROUP BY d.chapa, d.nome, g.materia, g.coddoc, d.DATAADMISSAO, d.CODSITUACAO, d.INSTRUCAOmec "
sqld=filtro
sqle="ORDER BY d.nome, g.materia, g.curso "
sql1=sqla & sqlb & sqlc & sqld & sqle
'response.write "<br>" & sql1
rs.Open sql1 , ,adOpenStatic , adLockReadOnly
'response.write rs.recordcount
inicio=1
'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1'
'	response.write "<td class="titulor">" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************

if rs.recordcount>0 then 
%>
<p class=realce>Relatório - Professores que ministram ou ministraram as Disciplinas:</p>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="grupo" align="center">Chapa</td>
	<td class="grupo" align="center">PROFESSOR</td>
	<td class="grupo" align="center">Titulação</td>
	<td class="grupo" align="center"><%=t(4)%></td>
	<td class="grupo" align="center"><%=t(3)%></td>
	<td class="grupo" align="center"><%=t(2)%></td>
	<td class="grupo" align="center"><%=t(1)%></td>
	<td class="grupo" align="center">Disciplina</td>
	<td class="grupo" align="center">Curso</td>
</tr>
<%
linhas=2
rs.movefirst
do while not rs.eof 
chapach=rs("chapa")
session("chapa")=chapach
if rs("materia")<>lastmateria then
end if

if linhas>65 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p style='margin-top:0; margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<p class=""titulor"">Relatório - Professores/Disciplinas"
	linhas=1
	response.write "<table border='1' cellpadding='0' cellspacing='1' style='border-collapse: collapse' width='690'>"
	response.write "<tr>"
	response.write "	<td class=""grupo"" align=""center"">Chapa</td>"
	response.write "	<td class=""grupo"" align=""center"">PROFESSOR</td>"
	response.write "	<td class=""grupo"" align=""center"">Titulação</td>"
	response.write "	<td class=""grupo"" align=""center"">" & t(4) & "</td>"
	response.write "	<td class=""grupo"" align=""center"">" & t(3) & "</td>"
	response.write "	<td class=""grupo"" align=""center"">" & t(2) & "</td>"
	response.write "	<td class=""grupo"" align=""center"">" & t(1) & "</td>"
	response.write "	<td class=""grupo"" align=""center"">Disciplina</td>"
	response.write "	<td class=""grupo"" align=""center"">Curso</td>"
	response.write "</tr>"
	linhas=linhas+1
end if
%>
<tr>
<%
if lastchapa=rs("chapa") then
	response.write "<td class=""campor"" colspan=7>&nbsp;</td>"
	response.write "<td class=""campor"">" & rs("materia") & "</td>"
	response.write "<td class=""campor"">" & rs("curso") & "</td>"
else
%>
	<td class="campoa"r style="border-top: 2px solid #000000" align="center"><%=rs("chapa")%></td>
	<td class="campoa"r style="border-top: 2px solid #000000"><%=rs("nome")%></td>
	<td class="campoa"r style="border-top: 2px solid #000000"><%=rs("instrucaomec")%></td>
	<td class="campoa"r style="border-top: 2px solid #000000" align="center"><%=rs("t1")%></td>
	<td class="campoa"r style="border-top: 2px solid #000000" align="center"><%=rs("t2")%></td>
	<td class="campoa"r style="border-top: 2px solid #000000" align="center"><%=rs("t3")%></td>
	<td class="campoa"r style="border-top: 2px solid #000000" align="center"><%=rs("t4")%></td>
	<td class="campor" style="border-top: 2px solid #000000"><%=rs("materia")%></td>
	<td class="campor" style="border-top: 2px solid #000000"><%=rs("curso")%></td>
<%
end if
%>
</tr>
<%
linhas=linhas+1
inicio=0
'lastmateria=rs("materia")
lastchapa=rs("chapa")

rs.movenext
loop
rs.close
set rs=nothing
%>
</table>
<p><i><font size="1" color="#0000FF"><b><%=selecao %></b></font></i></p>
<%	pagina=pagina+1
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
%>
<%
else 'sem registros
%>
<p>
<b><font color="#FF0000">
Esta seleção não mostra nenhum registro.</font></b></p>
<%
end if 'recordcount

end if 'if do request.form
conexao.close
set conexao=nothing
%>
</body>
</html>