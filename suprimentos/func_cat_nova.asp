<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a94")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Movimento de Estoque</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript" src="../date.js"></script>
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>
</head>
<body>
<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
		if request.form("bt_salvar")<>"" then
		tudook=1
		'if request.form("salvar")="1" then

if request.form("chapa")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione o funcionário!');</script>"
end if
if request.form("id_cat")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a categoria!');</script>"
end if

		sqla = "INSERT INTO uniforme_func_cat (id_cat, inicio, chapa, usuarioc, datac ) "
		sqlb = " SELECT " & request.form("id_cat") & ""
		sqlb=sqlb & ",'" & dtaccess(request.form("inicio")) & "' "
		sqlb=sqlb & ",'" & request.form("chapa") & "' "
		sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		sqlb=sqlb & ",getdate()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
		'end if
		end if 'request btsalvar
	else 'request.form=""
	end if

'if request.form("bt_salvar")="" then
if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then
%>
<form method="POST" action="func_cat_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Inclusão de Categoria de Funcionário</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
if request("item")<>"" then
	id_item=request("item")
elseif request.form("id_item")<>"" then
	id_item=request.form("id_item") 
else
	id_item=""
end if
if request("chapa")<>"" then
	chapa=request("chapa")
elseif request.form("chapa")<>"" then
	chapa=request.form("chapa") 
else
	chapa=""
end if
if request.form("id_cat")="" then id_cat="0" else id_cat=request.form("id_cat")
if request.form("inicio")="" then inicio=formatdatetime(now,2) else inicio=request.form("inicio")
'if request.form("chapa")="" then chapa="" else chapa=request.form("chapa")
%>
<!-- uniforme -->

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo colspan=2>Funcionário</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=chapa%>" name="chapa" size="5" onchange="chapa1()" onfocus="javascript:window.status='Informe o chapa do funcionário'"></td>
	<td class=fundo>
		<select size="1" name="nome" onchange="nome1()" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" >
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and codsindicato<>'03' "
if request("chapa")<>"" then sql2=sql2 & "and chapa='" & chapa & "' "
sql2=sql2 & "order by nome "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o funcionário....</option>"
rs.movefirst:do while not rs.eof
if chapa=rs("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rs("chapa")%>" <%=temp%>><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
		</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Categoria</td>
	<td class=titulo>A partir de</td>
</tr>
<tr>
	<td class=fundo>
		<select size="1" name="id_cat" onfocus="javascript:window.status='Selecione o tipo de uniforme'" >
<%
sql1="select id_cat, descricao from uniforme_categoria order by descricao "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione a categoria....</option>"
rs.movefirst:do while not rs.eof
if cstr(id_cat)=cstr(rs("id_cat")) then temp="selected" else temp=""
%>
	<option value="<%=rs("id_cat")%>" <%=temp%>><%=rs("descricao")%></option>
<%
rs.movenext:loop:rs.close
%>
	</select></td>
	<td class=fundo><input type="text" name="inicio" value="<%=inicio%>" size=10></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
	</td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2" onfocus="javascript:window.status='Clique para desfazer e limpar a tela'"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()" onfocus="javascript:window.status='Clique aqui para fechar sem salvar'"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing
end if   'request.form=""
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<!--
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

%>
</body>
</html>