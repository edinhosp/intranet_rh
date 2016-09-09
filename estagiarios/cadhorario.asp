<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")="N" or session("a72")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Bolsas de Estudo</title>
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
<body>
<%
registros=Session("RegistrosPorPagina")
registros=250
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")
conexao.open Application("conexao")
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("ordem")="" then session("ordem72")="Todos" else session("ordem72")=request.form("ordem")
	if request.form("localizar")="" then session("loc72")="" else session("loc72")=request.form("localizar")
	if isnumeric(session("loc72"))=true then session("loc72")=numzero(session("loc72"),5)

	if session("ordem72")="Todos" then
		texto="order by h.descricao "
	else
		texto="order by h.codigo "
	end if
	
	if session("loc72")<>"" then
   		if isnumeric(session("loc72")) then
			session("sql72a")="AND (h.codigo like '%" & session("loc72") & "%') "
		else
			session("sql72a")="AND (h.descricao like '%" & session("loc72") & "%') "
		end if
	else
		session("sql72a")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

sqla="select h.codigo, h.descricao, h.datacriacao, h.jsem, h.jmes, h.ativo " & _
"from est_cadhorario h " & _
"where h.codigo<>'' "
sqlb=" "
sqlc=texto 'sqlc="order by h.descricao "

sql1=sqla & sqlb & session("sql72a") & sqlc

if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	Set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez")="Nao"
else
	if request("folha")="" then pagina=1
	if request.form("pagina")<>"" then pagina=request.form("pagina")
	if request("folha")<>"" then pagina=request("folha")
	Session("Pagina")=pagina
	conexao.cursorlocation = 3 'aduseclient
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	Set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 	MostraDados
end if	

Sub MostraDados()
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
End Sub
%>

<form method="POST" name="form" action="cadhorario.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Horários para Estagiários</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="55%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""cadhorario.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""cadhorario.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
response.write "<a href=""cadhorario.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""cadhorario.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="30%" valign="top" align="center">
<% if session("a72")="T" then %>
<a href="cadhorario_nova.asp" onclick="NewWindow(this.href,'Inclusao','520','170','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif">
<font size="1">inserir novo horario</font></a>
<% end if %>

	</td>
	<td class=campo width="15%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="650" cellpadding="1" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Cod.</td>
	<td class=titulo align="center">Descrição</td>
	<td class=titulo align="center">desde</td>
	<td class=titulo align="center">J.Sem.</td>
	<td class=titulo align="center">&nbsp;</td>
	<td class=titulo align="center">&nbsp;</td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
jsemh=int(rs("jsem")/60)
jsemm=rs("jsem")-jsemh*60
%>
<tr>
	<td class=campo align="center"><%=rs("codigo")%></td>
	<td class=campo><%=rs("descricao") %></td>
	<td class=campo align="center"><%=rs("datacriacao")%></td>
	<td class=campo align="center"><%=jsemh&":"&numzero(jsemm,2)%></td>
	<td class=campo align="center">
	<% if session("a72")="T" or session("a72")="C" then %>
		<a href="cadhorario_alteracao.asp?codigo=<%=rs("codigo")%>" onclick="NewWindow(this.href,'AlteracaoHorario','520','170','no','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0" width=13 alt="Alterar os dados do horário"></a>
	<% end if %>
	</td>
	<td class=campo align="center">
	<% if session("a72")="T" or session("a72")="C" then %>
		<a href="cadhorario_dados.asp?codigo=<%=rs("codigo")%>" onclick="NewWindow(this.href,'QuadroHorario','600','400','yes','center');return false" onfocus="this.blur()">
		<img src="../images/Clock.gif" width="16" height="16" border="0" alt=""></a>
	<% end if %>
	</td>
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
<p><font color="red"><b>Esta seleção não mostra nenhum registro.</font></p>
<%
end if
%>
<hr>
<p style="margin-bottom:0;margin-top:5">
<font size=1>Ordem: <select size=1 name="ordem" class=a>
<option value="Todos" <%if session("ordem72")="Todos" then response.write "selected"%>>de descrição</option>
<option value="Tcod" <%if session("ordem72")="Tcod" then response.write "selected"%>>de código</option>
</select>

Localizar por descricao/codigo: <input type="text" name="localizar" size=35 value="<%=session("loc72")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
<br>
<input name="B2" type="submit" class="button" value="Clique para Filtrar">
</font></form>

</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>