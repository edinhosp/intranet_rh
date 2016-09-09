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
<title>Inclusão de Grupo de Rotinas</title>
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
		if request.form("nome_grupo")="" then
			tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe um nome para o grupo!');</script>"
		end if
		sql = "INSERT INTO rotinas_grupos (nome_grupo, usuarioc, datac ) "
		sql2 = " SELECT '" & request.form("nome_grupo") & "', " & _
		" '" & session("usuariomaster") & "', getdate() "
		sql1 = sql & sql2
		'response.write "<font size='1'>" & sql1
		if tudook=1 then conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
end if

'if request.form="" then
%>
<form method="POST" action="grupo_nova.asp">
<input type="hidden" name="id_grupo" size="4" value="0">
<table border="0" cellpadding="3" cellspacing="0" width="500">
	<tr><td class=grupo>Grupo de Rotinas</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Nome do Grupo de Rotinas</td>
</tr>
<tr>
	<td class=titulo>0</td>
	<td class=fundo><input type="text" name="nome_grupo" size="70" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing

'end if
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if

'	Response.write "<p>Registro salvo.<br>"
	'response.write '<script>javascript:top.window.close();</script>
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