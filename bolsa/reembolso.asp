<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a65")="N" or session("a65")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Reembolsos</title>
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
<body topmargin="5" leftmargin="5">
<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")
conexao.open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")

if request.form("B2")<>"" then
 	Session("PrimeiraVez65")="Sim"

	if request.form("localizar")="" then session("loc65")="" else session("loc65")=request.form("localizar")

	if isnumeric(session("loc65"))=true then session("loc65")=numzero(session("loc65"),5)

	if session("loc65")<>"" then
  		if isnumeric(session("loc65")) then
			session("sql65d")="AND (b.chapa like '%" & session("loc65") & "%') "
	   else
			session("sql65d")="AND (f.nome like '%" & session("loc65") & "%') "
		end if
	else
		session("sql65d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

sqla="SELECT b.id_bolsa, b.chapa, f.NOME, f.CODSITUACAO, b.tp_bolsa, b.nome_bolsista, b.curso, b.instituicao, b.observacao, Max(r.mes_base) AS ultima " & _
"FROM (bolsistas AS b INNER JOIN corporerm.dbo.pfunc AS f ON b.chapa = f.CHAPA collate database_default) LEFT JOIN bolsistas_reembolso AS r ON b.chapa = r.chapa " & _
"GROUP BY b.id_bolsa, b.chapa, f.NOME, f.CODSITUACAO, b.tp_bolsa, b.nome_bolsista, b.curso, b.instituicao, b.observacao " & _
"HAVING b.tp_bolsa in ('5','8','9','B','D') "
sqlc="ORDER BY f.NOME "
sql65=sqla & session("sql65d") & sqlc
	
if Session("PrimeiraVez65")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql65, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez65")="Nao"
else
	if request("folha")="" then pagina=1
	if request.form("pagina")<>"" then pagina=request.form("pagina")
	if request("folha")<>"" then pagina=request("folha")
	Session("Pagina")=pagina
	conexao.cursorlocation = 3 'aduseclient
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql65, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 	MostraDados
end if	
	
Sub MostraDados()
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
End Sub
%>
<form method="POST" name="form" action="reembolso.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Controle de Reembolsos</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="55%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""reembolso.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""reembolso.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
end if

response.write "&nbsp;<b>"
response.write "<select size='1' name='pagina' onchange='javascript:submit()'>"
for selpag=1 to rs.pagecount
	if selpag=atual then selpag1="selected" else selpag1=""
	response.write "<option value=" & selpag & " " & selpag1 & ">" & selpag & "</option>"
next
response.write "</select>"
response.write "/" & rs.pagecount & "</b>&nbsp;"

if atual=rs.pagecount or rs.pagecount=0 then
response.write "<img src='../images/setanext0.gif' border='0'>"
response.write "<img src='../images/setalast0.gif' border='0'>"
else
response.write "<a href=""reembolso.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""reembolso.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
    <td class=campo width="30%" valign="top" align="center">
<% if session("a65")="T" then %>
<!--
<a href="reembolso_nova.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoFolha','420','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo reembolso">
<font size="1">inserir novo lançamento</font></a>
-->
<% end if %>

	</td>
    <td class=campo width="15%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="690" cellpadding="1" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Sit.</td>
	<td class=titulo align="center">Curso</td>
	<td class=titulo align="center">Instituição</td>
	<td class=titulo align="center">Observação</td>
	<td class=titulo align="center">Ultimo</td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to Session("RegistrosPorPagina")
%>
<tr>
	<td class=campo align="center" <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    <a class=r href="reembolso_ver.asp?id_bolsa=<%=rs("id_bolsa")%>&chapa=<%=rs("chapa")%>&nome=<%=rs("nome")%>" onclick="NewWindow(this.href,'controlelanc','600','430','yes','center');return false" onfocus="this.blur()">
	<%=rs("chapa")%></a></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>> <%=rs("nome")%></td>
	<td class=campo align="center" <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>> <%=rs("codsituacao")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>> <%=rs("curso")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>> <%=rs("instituicao")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>> <%=rs("observacao")%></td>
	<td class=campo align="center" <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>> <%=month(rs("ultima")) & "/" & year(rs("ultima"))%></td>
</tr>
<%
if linha=1 then linha=0 else linha=1
rs.movenext
if rs.eof then exit for
'loop
Next
%>
</table>
<%
else 'sem registros
%>
<p><b><font color="#FF0000">Esta seleção não mostra nenhum registro.</font></b></p>
<%
end if 'sem registros
%>
<hr>

<br>
Localizar por nome/chapa: <input type="text" name="localizar" size=25 value="<%=session("loc65")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar"></p>
</form>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>