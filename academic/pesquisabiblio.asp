<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 1600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a93")="N" or session("a93")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Plano de Ensino - Relatório</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, rs3
set conexao=server.createobject ("ADODB.Connection")
conexao.open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

%>
<form method="POST" action="pesquisabiblio.asp" name="form" >

<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=450>
<tr>
	<td class=titulor nowrap>Tipo</td>
	<td class=titulor>Conteúdo</td>
	</tr>
<tr>
	<td class="campot"r nowrap><select size="1" name="selecao" <!--onChange="javascript:submit()"--> >
		<option value="LIVRE" <%if request.form("selecao")="LIVRE" then response.write "selected"%> >Livre</option>
		<option value="TITULO" <%if request.form("selecao")="TITULO" then response.write "selected"%> >Titulo</option>
		<option value="AUTOR" <%if request.form("selecao")="AUTOR" then response.write "selected"%> >Autor</option>
		</select>
	</td>
	<td class="campot"r>
		<input type="text" name="conteudo" size="30" value="<%=request.form("conteudo")%>">
		<input type="submit" name="b1" value="Pesquisar">
	</td>
</tr>
<tr>
	<td class=titulo></td>
	<td class=titulo>Referência</td>
</tr>
<%
if request.form("B1")<>"" then

sql1="select cod_acervo, referencia, classificacao, obra, ano_publicacao from pe_biblio where "
conteudo=request.form("conteudo")
select case request.form("selecao")
	case "LIVRE"
		sql2=" livre like '%" & conteudo & "%' or assunto like '%" & conteudo & "%' "
	case "TITULO"
		sql2=" titulo like '%" & conteudo & "%' "
	case "AUTOR"
		sql2=" autor like '%" & conteudo & "%' or autor_principal like '%" & conteudo & "'"
end select
sql=sql1 & sql2 & " order by obra, ano_publicacao "
inicio=now()

set rs3=server.createobject ("ADODB.Recordset")
set rs3.ActiveConnection = conexao
rs3.Open sql, ,adOpenStatic, adLockReadOnly

if rs3.recordcount>0 then
	do while not rs3.eof

'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse' width=350><tr>"
'for a= 0 to rs3.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs3.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'do while not rs3.eof 
'response.write "<tr>"
'for a= 0 to rs3.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs3.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs3.movenext
'loop
'response.write "</table><p>"
'*************** fim teste **********************
%>
<tr>
	<td class=campo style="border-bottom:2px solid #000000"><%=rs3("cod_acervo")%></td>
	<td class=campo style="border-bottom:2px solid #000000"><%=rs3("referencia")%></td>
</tr>
<%
	rs3.movenext
	loop
	termino=now():duracao=termino-inicio
	response.write "<tr><td class=grupo colspan=2>Pesquisou " & rs3.recordcount & " livros em " & cdbl(int(duracao*86400*100)/100) & " seg.</td></tr>"
else
	response.write "<tr><td class=grupo colspan=2>Nenhum registro encontrado</td></tr>"
end if 'rs3.recordcount
rs3.close


end if 'request.form
%>
</table>
</form>


</body>
</html>
<%
conexao.close
set conexao=nothing
%>