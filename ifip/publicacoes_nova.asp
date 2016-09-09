<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a30")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Publicações</title>
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
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql = "INSERT INTO ifip_publicacoes (id_ifip, dt_publicacao, titulo, [local], numero, pagina "
		sql = sql & ") "
		sql2 = " VALUES ( " & request.form("id_ifip") & ", "
		if request.form("dt_publicacao")="" then sql2=sql2 & "null, " else sql2=sql2 & " '" & dtaccess(request.form("dt_publicacao")) & "', "
		sql2=sql2 & " '" & request.form("titulo") & "', " 
		sql2=sql2 & " '" & request.form("local") & "', " 
		sql2=sql2 & " '" & request.form("numero") & "', " 
		sql2=sql2 & " '" & request.form("pagina") & "' " 
		sql1 = sql & sql2 & ")"
		'response.write "<font size='1'>" & sql1
		conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
	
end if

if request.form="" then
%>
<form method="POST" action="publicacoes_nova.asp" name="form">
<input type="hidden" name="id_ifip" size="4" value="<%=request("codigo")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Publicação</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Data da Publicação</td>
	<td class=titulo>Título</td>
</tr>
<tr>
	<td class=fundo valign=top><input type="text" name="dt_publicacao" size="8" value=""></td>
	<td class=fundo><textarea name="titulo" cols="50" rows="3"></textarea></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Local</td>
	<td class=titulo>Número</td>
	<td class=titulo>Página</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="local" size="20" value=""></td>
	<td class=fundo><input type="text" name="numero" size="6" value=""></td>
	<td class=fundo><input type="text" name="pagina" size="6" value=""></td>
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
else
'rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" then
	Response.write "<p>Registro salvo.<br>"
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>