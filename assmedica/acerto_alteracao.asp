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
	sql="UPDATE assmed_acertos SET "
	if request.form("data_acerto")<>"" then 
		sql=sql & "data_acerto = '"   & dtaccess(request.form("data_acerto"))  & "', "
	else
		sql=sql & "data_acerto = null, "
	end if
	sql=sql & "descricao    = '"  & request.form("descricao")    & "', "
	sql=sql & "empresa      = '"  & request.form("empresa")    & "', "
	sql=sql & "valor_acerto = "   & nraccess(request.form("valor_acerto")) & ", "
	sql=sql & "reembolso    = "   & nraccess(request.form("reembolso"))    & " "
	sql=sql & " WHERE id_acerto=" & session("id_alt_acerto")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM assmed_acertos WHERE id_acerto=" & session("id_alt_acerto")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_acerto=session("id_alt_acerto")
	else
		id_acerto=request("codigo")
	end if
	sqla="select * from assmed_acertos "
	sqlb="where id_acerto=" & id_acerto
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_acerto")=rs("id_acerto")

sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
%>
<form method="POST" action="acerto_alteracao.asp">
<input type="hidden" name="id_acerto" size="4" value="<%=rs("id_acerto")%>">  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Acerto de Nota Fiscal de Assistência Médica</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário Titular</td></tr>
<tr><td class=titulo><p class=realce><%=rs("chapa")%> - <%=rsnome("nome")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Data do Acerto</td>
	<td class=titulo>Descrição do Acerto</font></td></tr>
<tr><td class=fundo><input type="text" name="data_acerto" size="12" value="<%=rs("data_acerto")%>"></td>
	<td class=fundo><input type="text" name="descricao" size="45" value="<%=rs("descricao")%>"></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Empresa</td>
	<td class=titulo>Valor do Acerto</td>
	<td class=titulo>Reembolso</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="empresa">
<%
sqla="SELECT * from assmed_empresa ORDER by operadora"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rs("empresa")=rsc("codigo") then tempe="selected" else tempe=""
%>
		<option value="<%=rsc("codigo")%>" <%=tempe%>><%=rsc("operadora")%></option>
<%
rsc.movenext
loop
rsc.close
%>
		</select>
	</td>
	<td class=fundo><input type="text" name="valor_acerto" size="15" value="<%=rs("valor_acerto")%>"></td>
	<td class=fundo><input type="text" name="reembolso" size="15" value="<%=rs("reembolso")%>"></td>
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
set rsc=nothing
set rsnome=nothing
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