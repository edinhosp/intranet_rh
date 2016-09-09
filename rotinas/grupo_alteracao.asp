<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a92")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Alteração de Grupo de Rotinas</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
	tudook=1
if request.form("a1")="" and request.form("a2")="" and request.form("a3")="" and request.form("a4")="" and request.form("a5")="" and request.form("a6")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione os horários de aula!');</script>"
end if
		sql="UPDATE n_nomeacoes SET "
		sql=sql & "nomeacao = '"   & request.form("nomeacao")      & "', "
		sql=sql & "criacao  = '"   & request.form("criacao")     & "', "
		if request.form("extinta")="ON" then sql=sql & "extinta = 1 " else sql=sql & "extinta = 0 "
 			sql=sql & "usuarioa='" & session("usuariomaster") & "', "
			sql=sql & "dataa   =getdate(), "
		sql=sql & "WHERE id_nomeacao=" & session("id_alt_nomeacao")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM rotina_grupos WHERE id_grupo=" & session("id_alt_nomeacao")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_nomeacao=session("id_alt_nomeacao")
	else
		id_nomeacao=request("codigo")
	end if
	sqla="select * from n_nomeacoes "
	sqlb="where id_nomeacao=" & id_nomeacao
	sql1=sqla & sqlb 
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_nomeacao")=rs("id_nomeacao")
if rs("extinta")=0 then extinta="OFF" else extinta="ON"
if rs("extinta")=0 then extinta1="" else extinta1="checked"
%>
<form method="POST" action="tipo_alteracao.asp">
<input type="hidden" name="id_nomeacao" size="4" value="<%=rs("id_nomeacao")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Grupo de nomeações</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cód.</td>
	<td class="titulo">Nomeações para</td>
</tr>
<tr>
	<td class=titulo><%=rs("id_nomeacao")%></td>
	<td class=titulo><input type="text" name="nomeacao" size="70" value="<%=rs("nomeacao")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Criada em / pela</td>
	<td class=titulo>Extinta</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="criacao"  size="50" value="<%=rs("criacao")%>"></td>
	<td class=titulo><input type="checkbox" name="extinta" value="ON" <%=extinta1 %>></td>
	</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
	<input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser alterado!');</script>"
	end if
end if

conexao.close
set conexao=nothing
%>
</body>
</html>