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
<title>Inclusão de Acerto de NF</title>
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
	sql = "INSERT INTO assmed_acertos (" 
	sql = sql & "chapa, data_acerto, descricao, "
	sql = sql & "valor_acerto, empresa,reembolso "
	sql = sql & ") "
	sql2 = " SELECT "
	sql2=sql2 & " '" & request.form("chapa") & "', "
	if request.form("data_acerto")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("data_acerto")) & "', "
	sql2=sql2 & " '" & request.form("descricao") & "', "
	sql2=sql2 & " " & nraccess(request.form("valor_acerto")) & ", "
	sql2=sql2 & " '" & request.form("empresa") & "', "
	sql2=sql2 & " " & nraccess(request.form("reembolso")) & " "
	sql1 = sql & sql2 & ""
	'response.write "<font size='2'>" & sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
end if

if request.form("bt_salvar")="" or (request.form("bt_salvar")<>"" and tudook=0) then
%>
<form method="POST" action="acerto_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Acerto</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário Titular</td></tr>
<tr>
	<td class=fundo><select size="1" name="chapa" class=a>
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where chapa<'10000' "
if request("chapa")<>"" then sql2=sql2 & "and chapa='" & request("chapa") & "'" else sql2=sql2 & "order by nome"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if request("chapa")=rsc("chapa") then tempc="selected" else tempc=""
%>
	<option value="<%=rsc("chapa")%>" <%=tempc%>><%=rsc("nome")%></option>
<%
rsc.movenext
loop
rsc.close
%>
	</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Data do Acerto</td>
	<td class=titulo>Descrição do Acerto</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="data_acerto" size="12"></td>
	<td class=titulo><input type="text" name="descricao" size="45"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Empresa</td>
	<td class=titulo>Valor do Acerto</td>
	<td class=titulo>Reembolso</td>
</tr>
<tr>
	<td class=titulo>
	<select size="1" name="empresa">
<%
sqla="SELECT * from assmed_empresa ORDER by operadora"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
%>
	<option value="<%=rsc("codigo")%>"><%=rsc("operadora")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select>	  
	</td>
	<td class=titulo><input type="text" name="valor_acerto" size="15"></td>
	<td class=titulo><input type="text" name="reembolso" size="15"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar" class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
else
'rs.close
set rs=nothing
end if
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if

'	Response.write "<p>Registro salvo.<br>"
	'response.write '<script>javascript:top.window.close();</script>
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