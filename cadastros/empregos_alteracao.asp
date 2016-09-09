<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a88")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Outros empregos</title>
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
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql="UPDATE pfunc_empregos SET "
		sql=sql & "empresa = '" & request.form("empresa") & "', "
		sql=sql & "cargo   = '" & request.form("cargo")& "', "
		sql=sql & "desde   = '" & request.form("desde")   & "', "
		sql=sql & "ate     = '" & request.form("ate")   & "', "
		sql=sql & "usuarioe='" & session("usuariomaster") & "', "
		sql=sql & "datae   =getdate() "
		sql=sql & " WHERE id_emp=" & session("id_alt_emp")
		conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		sql="DELETE FROM pfunc_empregos WHERE id_emp=" & session("id_alt_emp")
		conexao.Execute sql, , adCmdText
	end if

else 'request.form=""

	if request("codigo")=null then
		id_form=session("id_alt_emp")
	else
		id_form=request("codigo")
	end if
	sqla="select * from pfunc_empregos "
	sqlb="where id_emp=" & id_form
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_emp")=rs("id_emp")

sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
%>
<form method="POST" action="empregos_alteracao.asp" name="form">
<input type="hidden" name="id_emp" size="4" value="<%=rs("id_emp")%>" style="font-size: 8 pt">
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Outros empregos</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=fundo><p class=realce><%=rs("chapa")%> - <%=rsnome("nome")%></p></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Empresa</td>
	<td class=titulo>Cargo</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="empresa" size="45" value="<%=rs("empresa")%>"></td>
	<td class=fundo><input type="text" name="cargo" size="30" value="<%=rs("cargo")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Desde:</td>
	<td class=titulo>Até:</td>
	<td class=titulo>&nbsp;</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="desde" size="10" value="<%=rs("desde")%>">  </td>
	<td class=fundo><input type="text" name="ate" size="10" value="<%=rs("ate")%>">  </td>
	<td class=fundo>&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
	<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
	<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
	<input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
	</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if
%>
<%
if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	Response.write "Registro atualizado.<br>"
	'response.write "<a href='javascript:window.close()'>Fechar Janela</a>"
%>
 <script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
<%
end if
%>
<%
set rsc=nothing
set rsnome=nothing
conexao.close
set conexao=nothing
%></body>
</html>