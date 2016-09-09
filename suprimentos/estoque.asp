<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a94")="N" or session("a94")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle Estoque Uniformes</title>
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
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B2")<>"" then
 	Session("PrimeiraVez94")="Sim"
	if request.form("localizar")="" then session("loc94e")="" else session("loc94e")=request.form("localizar")
	if isnumeric(session("loc94e"))=true then session("loc94e")=numzero(session("loc94e"),5)
	if session("loc94e")<>"" then
		session("sqld94e")="AND (i.descricao like '%" & session("loc94e") & "%') "
	else
		session("sqld94e")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
	registros=Session("RegistrosPorPagina")

	sqla="select i.id_item, i.descricao, i.tamanho, i.sequencia, i.codigorm, i.preco, i.qt_novo as novo, i.qt_usado as usado " & _
	"from uniforme_item i where i.id_item not in (select id_item from uniforme_estoque group by id_item) and (qt_usado+qt_novo)>0 " & _
	"union all " & _
	"SELECT e.id_item, i.descricao, i.tamanho, i.sequencia, i.codigoRM, i.preco, Sum(e.qt_novo*tipo) AS novo, Sum(e.qt_usado*tipo) AS usado " & _
	"FROM uniforme_tpmov t INNER JOIN (uniforme_item i INNER JOIN uniforme_estoque e ON i.id_item=e.id_item) ON t.id_mov=e.id_mov " & _
	"where e.id_item>0 "
	sqlc="GROUP BY e.id_item, i.descricao, i.tamanho, i.sequencia, i.codigoRM, i.preco "
	sqle="order by i.descricao, i.sequencia "
	sql36=sqla & session("sqld94e") & sqlc & sqle

	if Session("PrimeiraVez94")<>"Nao" then
		conexao.cursorlocation = 3 'aduseclient
		rs.CacheSize = registros
		rs.PageSize = registros
		set rs.ActiveConnection = conexao
		rs.Open sql36, ,adOpenStatic, adLockReadOnly
		Session("Pagina")=1
		MostraDados
		Session("PrimeiraVez94")="Nao"
	else
		if request("folha")="" then pagina=1
		if request.form("pagina")<>"" then pagina=request.form("pagina")
		if request("folha")<>"" then pagina=request("folha")
		Session("Pagina")=pagina
		conexao.cursorlocation = 3 'aduseclient
		rs.CacheSize = registros
		rs.PageSize = registros
		set rs.ActiveConnection = conexao
		rs.Open sql36, ,adOpenStatic, adLockReadOnly
		if rs.recordcount>0 then 	MostraDados
	end if	

	Sub MostraDados()
		if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
	End Sub
%>
<form method="POST" name="form" action="estoque.asp">
<p class="titulo" style="margin-top: 0; margin-bottom: 0">Movimentação Estoque de Uniformes</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class="campo" width="55%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""estoque.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""estoque.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
response.write "<a href=""estoque.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""estoque.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
    <td class="campo" width="30%" valign="top" align="center">
<% if session("a94")="T" then %>
<a href="estoque_nova.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoEstoque','420','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo lançamento" WIDTH="16" HEIGHT="16">
<font size="1">inserir novo lançamento</font></a>
<% end if %>

	</td>
    <td class="campo" width="15%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="2" width="650" cellpadding="0" style="border-collapse: collapse">
<tr>
	<td class="titulo" align="center" rowspan=2>Descrição Item</td>
	<td class="titulo" align="center" rowspan=2>Tamanho</td>
	<td class="titulo" align="center" colspan=2>Estoque</td>
</tr>
<tr>
	<td class="titulo" align="center">Novo</td>
	<td class="titulo" align="center">Usado</td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to Session("RegistrosPorPagina")
if linha=0 then classe="campo" else classe="campoa"
%>
<tr>
	<td class="<%=classe%>">
    <a class="r" href="estoque_ver.asp?item=<%=rs("id_item")%>" onclick="NewWindow(this.href,'controleFer','600','430','yes','center');return false" onfocus="this.blur()">
	&nbsp;<%=rs("Descricao")%></a></td>
	<td class="<%=classe%>" align="center"> <%=rs("tamanho")%> </td>
	<td class="<%=classe%>" align="center"> <%=rs("novo")%> </td>
	<td class="<%=classe%>" align="center"> <%=rs("usado")%> </td>
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
Localizar por descrição: <input type="text" name="localizar" size="25" value="<%=session("loc94e")%>">
Registros/Página: <input type="text" name="regpag" size="3" value="<%=Session("RegistrosPorPagina")%>">
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