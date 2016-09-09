<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a55")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Convênio com IES</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		sql="UPDATE rhconveniobe SET "
		sql=sql & "faculdade   = '"   & request.form("faculdade")   & "', "
		sql=sql & "mantenedora = '"   & request.form("mantenedora") & "', "
		sql=sql & "endereco    = '"   & request.form("endereco")    & "', "
		sql=sql & "cidade      = '"   & request.form("cidade")      & "', "
		sql=sql & "cnpj        = '"   & request.form("cnpj")        & "', "
		sql=sql & "email       = '"   & request.form("email")       & "', "
		sql=sql & "telefone    = '"   & request.form("telefone")    & "', "
		sql=sql & "contato     = '"   & request.form("contato")     & "'  "
		sql=sql & " WHERE id_faculdade=" & session("id_alt_faculdade")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM rhconveniobe WHERE id_faculdade=" & session("id_alt_faculdade")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null then
		id_faculdade=session("id_alt_faculdade")
	else
		id_faculdade=request("codigo")
	end if
	sqla="select * from rhconveniobe "
	sqlb="where id_faculdade=" & id_faculdade
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_faculdade")=rs("id_faculdade")
%>
<form method="POST" action="fac_alteracao.asp">
<input type="hidden" name="id_faculdade" size="4" value="<%=rs("id_faculdade")%>" >  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Convênio com IES</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Nome da Instituição</td></tr>
<tr>
	<td class=fundo><input type="text" name="faculdade"  size="60" value="<%=rs("faculdade")%>" ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Mantenedora</td>
	<td class=titulo>Telefone</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="mantenedora"  size="60" value="<%=rs("mantenedora")%>" ></td>
	<td class=fundo><input type="text" name="telefone"  size="20" value="<%=rs("telefone")%>" ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Endereço</td>
	<td class=titulo>Cidade</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="endereco"  size="50" value="<%=rs("endereco")%>" ></td>
	<td class=fundo><input type="text" name="cidade"  size="30" value="<%=rs("cidade")%>" ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500" height="54">
<tr>
	<td class=titulo>CNPJ</td>
	<td class=titulo>Contato</td>
	<td class=titulo>Email</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="cnpj"  size="25" value="<%=rs("cnpj")%>" ></td>
	<td class=fundo><input type="text" name="contato"  size="25" value="<%=rs("contato")%>" ></td>
	<td class=fundo><input type="text" name="email"  size="25" value="<%=rs("email")%>" ></td>
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

conexao.close
set conexao=nothing
%>
</body>
</html>