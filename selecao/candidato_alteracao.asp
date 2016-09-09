<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a68")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Alteração de Candidato</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="javascript" type="text/javascript"><!--
function nomeu() {
	form.nome_candidato.value=form.nome_candidato.value.toUpperCase()
}
// --></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

	if request.form("bt_salvar")<>"" then
		tudook=1
		sql="UPDATE rs_candidato SET "
		sql=sql & "nome_candidato='" & request.form("nome_candidato") & "', "
		sql=sql & "idade         = " & request.form("idade")          & " , "
		sql=sql & "email         ='" & request.form("email")          & "', "
		sql=sql & "telefone      ='" & request.form("telefone")       & "' "
		sql=sql & "WHERE id_candidato=" & session("id_alt_candidato")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM rs_candidato WHERE id_candidato=" & session("id_alt_candidato")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null then
		id_candidato=session("id_alt_candidato")
	else
		id_candidato=request("codigo")
	end if
	sqla="select * from rs_candidato "
	sqlb="where id_candidato=" & id_candidato
	sql1=sqla & sqlb 
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_candidato")=rs("id_candidato")
%>
<form method="POST" action="candidato_alteracao.asp">
<input type="hidden" name="id_candidato" size="4" value="<%=rs("id_candidato")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Alterar Candidato a vaga: <%=request("vaga")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Nome do Candidato</td>
</tr>
<tr>
	<td class=titulo><%=rs("id_candidato")%></td>
	<td class=titulo><input type="text" name="nome_candidato" size="70" value="<%=rs("nome_candidato")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Idade</td>
	<td class=titulo>Telefone</td>
	<td class=titulo>E-Mail</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="idade"  size="3" value="<%=rs("idade")%>"></td>
	<td class=titulo><input type="text" name="telefone" size="15" value="<%=rs("telefone")%>"></td>
	<td class=titulo><input type="text" name="email" size="40" value="<%=rs("email")%>"></td>
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
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser alterado!');</script>"
	end if
	'Response.write "<p>Registro atualizado.<br>"
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<!--
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if
%>
</body>
</html>