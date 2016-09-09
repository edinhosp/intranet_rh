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
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Alteração de Histórico</title>
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
		sql="UPDATE ifip_historico SET "
		if request.form("dt_historico")=""     then
			sql=sql & "dt_historico=null,"
		else
			sql=sql & "dt_historico='" & dtaccess(request.form("dt_historico")) & "', "
		end if
		sql=sql & "sequencia =" & request.form("sequencia") & ", "
		sql=sql & "historico='" & request.form("historico") & "', "
		sql=sql & "observacao='" & request.form("observacao") & "' "
		sql=sql & "WHERE id_hist=" & session("id_alt_hist")
		conexao.Execute sql ', , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		sql="DELETE FROM ifip_historico WHERE id_hist=" & session("id_alt_hist")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
	if request("codigo")=null then
		id_hist=session("id_alt_hist")
	else
		id_hist=request("codigo")
	end if
	sqla="select * from ifip_historico "
	sqlb="where id_hist=" & id_hist
	sql1=sqla & sqlb 
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_hist")=rs("id_hist")
%>
<form method="POST" action="historico_alteracao.asp" name="form">
<input type="hidden" name="id_hist" size="4" value="<%=rs("id_hist")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Alteração de Históricos</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Sequência</td>
	<td class=titulo>Data</td>
	<td class=titulo>Histórico</td>
</tr>
<tr>
	<td class=fundo valign=top><%=rs("sequencia")%><input type="hidden" name="sequencia" value="<%=rs("sequencia")%>"></td>
	<td class=fundo valign=top><input type="text" name="dt_historico" size="8" value="<%=rs("dt_historico")%>"></td>
	<td class=fundo><textarea name="historico" cols="45" rows="3"><%=rs("historico")%></textarea></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Observação</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="observacao" size="70" value="<%=rs("observacao")%>"></td>
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