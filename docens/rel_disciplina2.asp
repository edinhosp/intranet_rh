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
<title>Pesquisa de Formação</title>
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
<form method="POST" action="rel_disciplina2.asp" name="form">
  <p><font color="#0000ff"><b>Seleções para o relatório &quot;Pesquisa de Formação Acadêmica&quot;</b></font></p>
  <p>Palavras chaves nas cursos e/ou abrangência a pesquisar:<br>
  palavra <input type="text" name="T1" size="20" class="form_box"><br>
  	<input type="radio" name="tipo" value="01">Administrativos
  	<input type="radio" name="tipo" value="03" checked>Professores
  	<input type="radio" name="tipo" value="99">Todos<br>
  
<!--
  palavra 2 <input type="text" name="T2" size="20" class="form_box"><br>
  palavra 3 <input type="text" name="T3" size="20" class="form_box"><br>
  palavra 4 <input type="text" name="T4" size="20" class="form_box"><br>
  palavra 5 <input type="text" name="T5" size="20" class="form_box"><br>
  palavra 6 <input type="text" name="T6" size="20" class="form_box"></p>-->
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
filtro="":filtro2="":selecao=""
chave=0
palavra=request.form("T1")
'for a=1 to 6
'	palavra=request.form("T" & a)
'	if palavra<>"" then
'		if chave=0 then filtro="HAVING ":selecao="Seleção: disciplinas relacionadas com "
'		if chave=1 then filtro=filtro & " or ":selecao=selecao & ", "
'		filtro=filtro & "g.materia Like '%" & palavra & "%' " : selecao=selecao & palavra
'		if chave=0 then chave=1
'	end if
'next

if request.form("tipo")="01" then tiposql=" and f.codsindicato<>'03' "
if request.form("tipo")="03" then tiposql=" and f.codsindicato='03' "
if request.form("tipo")="99" then tiposql=""
sql1="select c.codprof as chapa, f.nome, c.tipo2 as tipo, c.curso, a.descricao as abrangencia " & _
"from uprofformacao_ c, dc_professor f, uprof_abrangencia a " & _
"where c.codprof=f.chapa collate database_default and f.codsituacao in ('A','F','Z','E') and c.abrangencia=a.abrangencia " & _
"and (c.curso like '%" & palavra & "%' or a.descricao like '%" & palavra & "%') " & tiposql
'response.write sql1
rs.Open sql1 , ,adOpenStatic , adLockReadOnly

inicio=1
'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
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
<p class=realce>Professores com cursos ou cursos com abrangência: <%=palavra%></p>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="grupo" align="center">Chapa</td>
	<td class="grupo" align="center">Professor</td>
	<td class="grupo" align="center">Tipo</td>
	<td class="grupo" align="center">Curso</td>
	<td class="grupo" align="center">Área de abrangência</td>
</tr>
<%
linhas=2
rs.movefirst
do while not rs.eof 
chapach=rs("chapa")
%>
<tr>
	<td class=campo style="border-top: 2px solid #000000" align="center"><%=rs("chapa")%></td>
	<td class=campo style="border-top: 2px solid #000000"><%=rs("nome")%></td>
	<td class=campo style="border-top: 2px solid #000000"><%=rs("tipo")%></td>
	<td class=campo style="border-top: 2px solid #000000"><%=rs("curso")%></td>
	<td class=campo style="border-top: 2px solid #000000"><%=rs("abrangencia")%></td>
</tr>
<%
linhas=linhas+1
inicio=0
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