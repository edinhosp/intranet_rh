<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
if session("a4")="N" or session("a4")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Configuração e Preferências</title>
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->
<%
'se for alterar registro por cookie
'response.cookies("vrh06")("registropagina")="25"

'para ler o cookie
rp=request.cookies("vrh06")("registropagina")

dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	if request.form("novasenha")<>request.form("novasenha2") then
		tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('A senha digitada não confere!');</script>"
	end if
	if request.form("novasenha")="" or request.form("novasenha2")="" then
		tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('A senha deve ser digitada!');</script>"
	end if 
	if tudook=1 then
		sql="update usuarios set password='" & request.form("novasenha") & "' where usuario='" & session("usuariomaster") & "' "
		conexao.execute sql
		response.write "<script language='JavaScript' type='text/javascript'>alert('Senha alterada!');</script>"
	end if
end if

sql="SELECT nome, usuario, password, estilo, timeout, new " & _
"FROM usuarios WHERE usuario='" & session("usuariomaster") & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" action="preferencias.asp" name="form">
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse">
<tr><td class="grupo" colspan="3">Preferências e configurações</td></tr>
<tr><td class="titulo">Nome</td>
	<td class="campo"><%=rs("nome")%></td>
	<td class="campo"></tr>
<tr><td class="titulo">Código</td>
	<td class="campo"><%=rs("usuario")%></td>
	<td class="campo"></tr>
<tr><td class="titulo">Senha</td>
	<td class="campo">*****<%password=rs("password")%></td>
	<td class="campo">Nova senha: <input type="text" name="novasenha" size=6>
	Confirmar senha: <input type="text" name="novasenha2" size=6>
	</tr>
<tr><td class="titulo">Tempo de inatividade</td>
	<td class="campo"><%=rs("timeout")%> minutos</td>
	<td class="campo"></tr>
<tr><td class="titulo">Registro por páginas</td>
	<td class="campo"><%=session("registrosporpagina")%></td>
	<td class="campo"><%=rp%></tr>
</table>
<br>
<input type="submit" value="Salvar Alteração" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
</form>
<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>


<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>