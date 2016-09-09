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
<title>Inclusão de Convênio com IES</title>
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

if request.form("bt_salvar")<>"" then
	tudook=1

if request.form("faculdade")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o nome da Instituição de Ensino!');</script>"

	sql = "INSERT INTO rhconveniobe (" 
	sql = sql & "faculdade, mantenedora, endereco, cidade, "
	sql = sql & "cnpj, contato, email, telefone "
	sql = sql & ") "
	sql2 = " SELECT "
	sql2=sql2 & " '" & request.form("faculdade") & "', "
	sql2=sql2 & " '" & request.form("mantenedora") & "', "
	sql2=sql2 & " '" & request.form("endereco") & "', "
	sql2=sql2 & " '" & request.form("cidade") & "', "
	sql2=sql2 & " '" & request.form("cnpj") & "', "
	sql2=sql2 & " '" & request.form("contato") & "', "
	sql2=sql2 & " '" & request.form("email") & "', "
	sql2=sql2 & " '" & request.form("telefone") & "' "
	sql1 = sql & sql2 & ""
	'response.write "<font size='1'>" & sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
end if

%>
<form method="POST" action="fac_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Convênio com IES</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Nome da Instituição</td></tr>
<tr><td class=fundo><input type="text" name="faculdade" size="60" value="<%=request.form("faculdade")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Mantenedora</td>
	<td class=titulo>Telefone</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="mantenedora" size="60" value="<%=request.form("mantenedora")%>"></td>
	<td class=fundo><input type="text" name="telefone" size="20" value="<%=request.form("telefone")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Endereço</td>
	<td class=titulo>Cidade</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="endereco" size="50" value="<%=request.form("endereco")%>"></td>
	<td class=fundo><input type="text" name="cidade" size="30" value="<%=request.form("cidade")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>CNPJ</td>
	<td class=titulo>Contato</td>
	<td class=titulo>Email</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="cnpj" size="25" value="<%=request.form("cnpj")%>">    </td>
	<td class=fundo><input type="text" name="contato" size="25" value="<%=request.form("contato")%>">    </td>
	<td class=fundo><input type="text" name="email" size="25" value="<%=request.form("email")%>">    </td>
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