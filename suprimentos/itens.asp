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
<title>Uniformes</title>
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
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function periodo1() {	form.perlanc.value="Todos";	}
function perlanc1() {	form.periodo.value="Todos";	}
--></script>
</head>
<body>
<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rsc=server.createobject ("ADODB.Recordset")
set rsc.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request("codigo")<>"" then session("cat94")=request("codigo"):session("sqla94")="AND l.id_cat=" & session("cat94") & " "

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("categoria")="" then session("cat94")="Todas" else session("cat94")=request.form("categoria")
	if session("cat94")<>"Todas" then
		session("sqla94")="AND l.id_cat=" & session("cat94") & " "
	else
		session("sqla94")=""
	end if

	if request.form("tamanho")="" then session("tam94")="Todas" else session("tam94")=request.form("tamanho")
	if session("tam94")<>"Todas" then
		session("sqlb94")="AND i.tamanho='" & session("tam94") & "' "
	else
		session("sqlb94")=""
	end if

	if request.form("localizar")="" then session("loc94i")="" else session("loc94i")=request.form("localizar")
	if session("loc94i")<>"" then
		session("sqlc94i")="AND i.descricao like '%" & session("loc94i") & "%' "
	else
		session("sqlc94i")=""
	end if
		
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

sqla="select i.id_item, i.descricao, i.codigorm, i.tamanho, i.sequencia, i.qt_novo, i.qt_usado, i.preco " & _
"from uniforme_item i, uniforme_link l WHERE i.id_item=l.id_item "
sqlb=""
sqlc="GROUP BY i.id_item, i.descricao, i.codigorm, i.tamanho, i.sequencia, i.qt_novo, i.qt_usado, i.preco "
sqlc=sqlc & "ORDER BY i.descricao, i.sequencia, i.tamanho "

sql1=sqla & sqlb & session("sqla94") & session("sqlb94") & session("sqlc94i") & sqlc
if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	'conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
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
	'conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 	MostraDados
end if	

Sub MostraDados()
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
End Sub
%>
<form method="POST" action="itens.asp" name="form">
<input type="hidden" name="vez1" value="<%=session("PrimeiraVez")%>">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Uniformes</p>
<table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo width="60%" valign="center" align="left">Página: 
<%
Session("Load1")="1"
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""itens.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""itens.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
end if

response.write "&nbsp;<b>"
response.write "<select size='1' name='pagina' onChange='javascript:submit()'>"
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
response.write "<a href=""itens.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""itens.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="20%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
	<td class=campo width="20%" valign="top" align="right">
<% if session("a94")="T" then %>
<a href="itens_nova.asp" onclick="NewWindow(this.href,'InclusaoUniforme','530','330','no','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
<font size="1">inserir novo uniforme</font></a>
<% end if %>
	</td>
</tr>
</table>

<table border="1" cellspacing="0" cellpadding="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulor align="center" rowspan=2 width="30%">Descrição</td>
	<td class=titulor align="center" rowspan=2>Código RM</td>
	<td class=titulor align="center" rowspan=2>Tamanho</td>
	<td class=titulor align="center" rowspan=2 width=100>Categorias</td>
	<td class=titulor align="center" colspan=2>Estoque Inicial</td>
	<td class=titulor align="center" rowspan=2>Preço</td>
	<td class=titulor align="center" rowspan=2><img border="0" src="../images/Magnify.gif"></td>
</tr>
<tr>
	<td class=titulor align="center">Qt.Novo</td>
	<td class=titulor align="center">Qt.Usado</td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
	<td class=campo nowrap><%=rs("descricao")%></td>
	<td class=campo><%=rs("codigorm")%></td>
	<td class=campo align="center"><%=rs("tamanho")%></td>
	<td class="campor" align="left">
<%	
sqlc="select c.descricao from uniforme_categoria c, uniforme_link l where c.id_cat=l.id_cat and l.id_item=" & rs("id_item")
rs2.Open sqlc, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	rs2.movefirst:do while not rs2.eof
	'if rs2.absoluteposition>1 and rs2.absoluteposition<=rs2.recordcount then response.write "<br>"
	if rs2.absoluteposition>1 and rs2.absoluteposition<=rs2.recordcount then response.write ", "
	response.write rs2("descricao")
	rs2.movenext:loop
end if
rs2.close	
%>	
	</td>
	<td class=campo align="center"><%=rs("qt_novo")%></td>
	<td class=campo align="center"><%=rs("qt_usado")%></td>
	<td class=campo align="center"><%=formatnumber(rs("preco"),2)%></td>
	<td class=campo align="center">
    <% if session("a94")="T" then %>
		<a href="itens_alteracao.asp?codigo=<%=rs("id_item")%>" onclick="NewWindow(this.href,'AlteracaoUniforme','530','330','no','center');return false" onfocus="this.blur()">
		<img border='0' src='../images/folder95o.gif'></a>
	<% end if %>
	</td>
</tr>
<%
rs.movenext
if rs.eof then exit for
'loop
Next

else 'sem registros
%>
<td class=grupo colspan=13>Esta seleção não mostra nenhum registro.</td>
<%
end if
%>
</table>

<p><font size="1">
<%
sql2="select id_cat, descricao from uniforme_categoria order by descricao"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Categoria: <select size="1" name="categoria">
<option value="Todas" <%if session("cat94")="Todas" then response.write "selected"%>>Todas categorias</option>
<%
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("id_cat")%>" <%if cstr(session("cat94"))=cstr(rs2("id_cat")) then response.write "selected"%>><%=rs2("descricao")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<!-- turmas -->
&nbsp;&nbsp;Tamanho: <select size="1" name="tamanho">
<option value="Todas" <%if session("tam94")="Todas" then response.write "selected"%>>Todos</option>
<%
sql2="SELECT tamanho from uniforme_item group by tamanho"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("tamanho")%>" <%if session("tam94")=rs2("tamanho") then response.write "selected"%>><%=rs2("tamanho")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<br>
Localizar por descrição: <input type="text" name="localizar" size=35 value="<%=session("loc94i")%>">

Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar">
</form>
</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>