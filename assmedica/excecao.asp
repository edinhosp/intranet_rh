<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a81")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Acerto de NF</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(10)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	sql="UPDATE assmed_beneficiario SET "
	if request.form("dt_canc")<>"" then 
		sql=sql & "dt_canc='" & dtaccess(request.form("dt_canc"))  & "', "
	else
		sql=sql & "dt_canc=null, "
	end if
	sql=sql & " obs='" & request.form("obs") & "' "
	sql=sql & " WHERE chapa='" & session("id_alt_excecao") & "' "
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_excecao=session("id_alt_excecao")
	else
		id_excecao=request("codigo")
	end if
	sqla="select * from assmed_beneficiario "
	sqlb="where chapa='" & id_excecao & "' "
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_excecao")=rs("chapa")

%>
<form method="POST" action="excecao.asp">
<input type="hidden" name="chapa" size="4" value="<%=rs("chapa")%>">  
<table border="0" cellpadding="3" cellspacing="0" width="480">
<tr><td class=grupo>Exceção de Permanência</td></tr>
</table>


<table border="0" cellpadding="3" cellspacing="0" width="480">
<tr><td class=titulo>Data Limite</td>
	<td class=titulo>Motivo</font></td></tr>
<tr><td class=fundo><input type="text" name="dt_canc" size="12" value="<%=rs("dt_canc")%>"></td>
	<td class=fundo><input type="text" name="obs" size="45" value="<%=rs("obs")%>"></td></tr>
</table>


<table border="0" cellpadding="3" cellspacing="0" width="480">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if
set rsc=nothing
set rsnome=nothing
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
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