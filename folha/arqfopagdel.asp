<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")="N" or session("a48")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Exclusão de Parcela</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1() { form.nome.value=form.nome.value.toUpperCase() }
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

dataparc=request("dataparc")
nroperiodo=request("nroperiodo")
nroparcela=request("nroparcela")
dtpagto=request("dtpagto")
response.write "<br>" & dataparc
response.write "<br>" & nroperiodo
response.write "<br>" & nroparcela
response.write "<br>" & dtpagto

if request.form<>"" then
	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM creditofolhaparcelas where dataparc='" & dtaccess(request.form("dataparc")) & "' " & _
		"and nroperiodo=" & request.form("nroperiodo") & " and nroparcela=" & request.form("nroparcela") & " and dtpagto='" & dtaccess(dtpagto) & "' "
		response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

%>
<form method="POST" action="arqfopagdel.asp" name="form">
<input type="hidden" name="dataparc" value="<%=dataparc%>">
<input type="hidden" name="nroperiodo" value="<%=nroperiodo%>">
<input type="hidden" name="nroparcela" value="<%=nroparcela%>">
<input type="hidden" name="dtpagto" value="<%=dtpagto%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
	<tr><td class="grupo">Confirmação de exclusão de Parcela - Folha devida em <%=dtpagto%></td></tr>
	<tr><td class="campo">Deseja excluir a parcela <%=nroparcela%> do período <%=nroperiodo%> da data de pagamento <%=dataparc%>?</td></tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
<tr>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%

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