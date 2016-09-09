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
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Alteração de Categoria de Uniforme</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql="UPDATE uniforme_categoria SET "
		sql=sql & "descricao = '" & request.form("descricao") & "' "
		sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		sql=sql & ",dataa   =getdate() "
		sql=sql & "WHERE id_cat=" & session("id_alt_cat")
		conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		sql="DELETE FROM uniforme_categoria WHERE id_cat=" & session("id_alt_cat")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
	if request("codigo")=null then
		id_cat=session("id_alt_cat")
	else
		id_cat=request("codigo")
	end if
	sqla="select * from uniforme_categoria where id_cat=" & id_cat
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
end if

if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_cat")=rs("id_cat")
%>
<form method="POST" action="categ_alteracao.asp">
<input type="hidden" name="id_cat" size="4" value="<%=rs("id_cat")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Categoria de Uniforme</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cód.</td>
	<td class="titulo">Descrição</td>
</tr>
<tr>
	<td class=titulo><%=rs("id_cat")%></td>
	<td class=titulo><input type="text" name="descricao" size="70" value="<%=rs("descricao")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
	<!--<input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>-->
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	Response.write "<p>Registro atualizado."
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar janela" class="button" onClick="top.window.close()">
</form>
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>