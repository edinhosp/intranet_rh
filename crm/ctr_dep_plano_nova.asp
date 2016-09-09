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
<title>Inclusão de Assistência Médica</title>
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
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	if request.form("compr")="ON" then compr = -1 else compr = 0
	sql = "INSERT INTO assmed_dep_mudanca (" 
	sql = sql & "chapa, nrodepend, empresa, plano, codigo, "
	sql = sql & "inclusao,ivigencia, fvigencia, compr, oper, uoper"
	sql = sql & ") "
	sql2 = " SELECT "
	sql2=sql2 & " '" & request.form("chapa") & "', "
	sql2=sql2 & " " & request.form("nrodepend") & ", "
	sql2=sql2 & " '" & request.form("empresa") & "', "
	sql2=sql2 & " '" & request.form("plano") & "', "
	sql2=sql2 & " '" & request.form("codigo1") & "', "
	if request.form("inclusao")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("inclusao")) & "', "
	if request.form("ivigencia")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("ivigencia")) & "', "
	if request.form("fvigencia")="" then sql2=sql2 & "'12/31/2020'," else sql2=sql2 & " '" & dtaccess(request.form("fvigencia")) & "', "
	sql2=sql2 & compr & ", "
	sql2=sql2 & " '" & request.form("oper") & "', "
	sql2=sql2 & " '" & request.form("uoper") & "' "
	sql1 = sql & sql2 & ""
	'response.write "<font size='2'>" & sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
end if

if request.form("bt_salvar")="" or (request.form("bt_salvar")<>"" and tudook=0) then
if request("chapa")="" then chapa=request.form("chapa") else chapa=request("chapa")
if request("nrodepend")="" then nrodepend=request.form("nrodepend") else nrodepend=request("nrodepend")
%>
<form method="POST" action="ctr_dep_plano_nova.asp" name="form" >
<input type="hidden" name="chapa" value="<%=chapa%>">
<input type="hidden" name="nrodepend" value="<%=nrodepend%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Assistência Médica</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Nome do Dependente</td></tr>
<tr><td class=titulo><select size="1" name="id_dep" class=a>
<%
sql2="select nome from corporerm.dbo.pfdepend where chapa='" & chapa & "' and nrodepend=" & nrodepend
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
'if codigo=rsc("id_dep") then tempc="selected" else tempc=""
%>
	<option value="" <%=tempc%>><%=rsc("nome")%></option>
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
	<td class=titulo>Empresa de saúde</td>
	<td class=titulo>Plano escolhido</td>
</tr>
<tr>
	<td class=titulo><select size="1" name="empresa" onchange="javascript:submit()">
<%
if request.form("empresa")="" then empresa=0 else empresa=request.form("empresa")
sqla="SELECT * from assmed_empresa ORDER by operadora"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if empresa=rsc("codigo") then tempt="selected" else tempt=""
%>
	<option value="<%=rsc("codigo")%>" <%=tempt%>><%=rsc("operadora")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select></td>
	<td class=titulo><select size="1" name="plano">
		<option value="">Selecione um plano de saúde</option>
<%
sqla="SELECT * from assmed_planos where codigo='" & empresa & "' ORDER by seq, plano"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
rsc.movefirst
do while not rsc.eof
if request.form("plano")=rsc("plano") then tempp="selected" else tempp=""
%>
	<option value="<%=rsc("plano")%>" <%=tempp%>><%=rsc("plano") %></option>
<%
rsc.movenext:loop
end if
rsc.close
%>
	</select>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Código da carteirinha</td>
	<td class=titulo>Inclusão</td>
	<td class=titulo>Início Cobr.</td>
	<td class=titulo>Término</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="codigo1"   size="25" value="<%=request.form("codigo1")%>"></td>
	<td class=titulo><input type="text" name="inclusao" size="12" value="<%=request.form("inclusao")%>"></td>
	<td class=titulo><input type="text" name="ivigencia" size="12" value="<%=request.form("ivigencia")%>"></td>
	<td class=titulo><input type="text" name="fvigencia" size="12" value="<%=request.form("fvigencia")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Operação atual de Cadastro</td>
	<td class=titulo>Ultima operação</td>
	<td class=titulo>Emitiu Comprovante</td>
</tr>
<tr>
	<td class=titulo><select size="1" name="oper">
		<option value=""></option>
		<option value="I" <%if request.form("oper")="I" then response.write "selected"%>>Inclusão</option>
		<option value="A" <%if request.form("oper")="A" then response.write "selected"%>>Alteração</option>
		<option value="E" <%if request.form("oper")="E" then response.write "selected"%>>Exclusão</option>
		<option value="2" <%if request.form("oper")="2" then response.write "selected"%>>2ª Via</option>
		</select></td>
	<td class=titulo><select size="1" name="uoper">
		<option value=""></option>
		<option value="I" <%if request.form("uoper")="I" then response.write "selected"%>>Inclusão</option>
		<option value="A" <%if request.form("uoper")="A" then response.write "selected"%>>Alteração</option>
		<option value="E" <%if request.form("uoper")="E" then response.write "selected"%>>Exclusão</option>
		<option value="2" <%if request.form("uoper")="2" then response.write "selected"%>>2ª Via</option>
		</select></td>
	<td class=titulo><input type="checkbox" name="compr" value="ON" <%if request.form("compr")="ON" then response.write "checked"%>></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="  Fechar  " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
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