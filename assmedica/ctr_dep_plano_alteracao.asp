<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a81")="N" or session("a81")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Altera��o de Assist�ncia M�dica do dependente</title>
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
	sql="UPDATE assmed_dep_mudanca SET "
	sql=sql & "empresa  = '"   & request.form("empresa") & "', "
	sql=sql & "plano    = '"   & request.form("plano")   & "', "
	sql=sql & "codigo   = '"   & request.form("codigo1")  & "', "
	if request.form("inclusao")<>"" then 
		sql=sql & "inclusao = '"   & dtaccess(request.form("inclusao")) & "', "
	else
		sql=sql & "inclusao = null, "
	end if
	if request.form("ivigencia")<>"" then 
		sql=sql & "ivigencia = '"   & dtaccess(request.form("ivigencia")) & "', "
	else
		sql=sql & "ivigencia = null, "
	end if
	if request.form("fvigencia")<>"" then 
		sql=sql & "fvigencia = '"   & dtaccess(request.form("fvigencia")) & "', "
	else
		sql=sql & "fvigencia = null, "
	end if
	if request.form("compr")="ON" then 
		sql=sql & "compr = 1, " 
	else
		sql=sql & "compr = 0, "
	end if
	sql=sql & "oper      = '"   & request.form("oper")   & "', "
	sql=sql & "uoper     = '"   & request.form("uoper")  & "' "
	sql=sql & " WHERE id_mud=" & session("id_alt_mud")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM assmed_dep_mudanca WHERE id_mud=" & session("id_alt_mud")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_mud=session("id_alt_mud")
	else
		id_mud=request("codigo")
	end if
	sqla="select * from assmed_dep_mudanca "
	sqlb="where id_mud=" & id_mud
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_mud")=rs("id_mud")

sqlz="select nome from corporerm.dbo.pfdepend where chapa='" & rs("chapa") & "' and nrodepend=" & rs("nrodepend") & ""
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
%>
<form method="POST" action="ctr_dep_plano_alteracao.asp" name="planodesaude">
<input type="hidden" name="id_mud" size="4" value="<%=rs("id_mud")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Assist�ncia M�dica do dependente</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Nome do Dependente</td></tr>
<tr><td class=titulo><p class=realce><%=rs("nrodepend")%> - <%=rsnome("nome")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Empresa de sa�de</td>
	<td class=titulo>Plano escolhido</td>
</tr>
<tr>
	<td class=titulo><select size="1" name="empresa" onchange="javascript:submit()">
<%
if request.form("empresa")="" then empresa=rs("empresa") else empresa=request.form("empresa")
sqla="SELECT * from assmed_empresa ORDER by operadora"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if rsc("codigo")=empresa then tempt="selected" else tempt=""
%>
	<option value="<%=rsc("codigo")%>" <%=tempt%>><%=rsc("operadora")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select></td>
	<td class=titulo><select size="1" name="plano">
		<option value="">Selecione um plano de sa�de</option>
<%
if request.form("plano")="" then plano=rs("plano") else plano=request.form("plano")
sqla="SELECT * from assmed_planos where codigo='" & rs("empresa") & "' ORDER by seq, plano"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if rsc("plano")=plano then tempp="selected" else tempp=""
%>
	<option value="<%=rsc("plano")%>" <%=tempp%>><%=rsc("plano") %></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>C�digo da carteirinha</td>
	<td class=titulo>Inclus�o</td>
	<td class=titulo>In�cio Cobr.</td>
	<td class=titulo>T�rmino</td>
</tr>
<tr>
<%
if request.form("codigo1")=""    then codigo1  =rs("codigo")    else codigo1   =request.form("codigo1")
if request.form("inclusao")="" then inclusao=rs("inclusao") else inclusao=request.form("inclusao")
if request.form("ivigencia")="" then ivigencia=rs("ivigencia") else ivigencia=request.form("ivigencia")
if request.form("fvigencia")="" then fvigencia=rs("fvigencia") else fvigencia=request.form("fvigencia")
if len(codigo1)=16 then tamanho="" else tamanho="Cod.Inc."
%>
	<td class=titulo><input type="text" name="codigo1"  size="16" maxlenght=16 value="<%=codigo1%>"><%=tamanho%></td>
	<td class=titulo><input type="text" name="inclusao" size="12" value="<%=inclusao%>"></td>
	<td class=titulo><input type="text" name="ivigencia"  size="12" value="<%=ivigencia%>"></td>
	<td class=titulo><input type="text" name="fvigencia"  size="12" value="<%=fvigencia%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Opera��o atual de Cadastro</td>
	<td class=titulo>Ultima opera��o</td>
	<td class=titulo>Emitiu Comprovante</td>
</tr>
<tr>
<%
if request.form("oper")=""  then oper =rs("oper")  else oper =request.form("oper")
if request.form("uoper")="" then uoper=rs("uoper") else uoper=request.form("uoper")
if rs("compr")=0 and request.form("compr")<>"ON" then obs1="" else obs1="checked"
%>
	<td class=titulo><select size="1" name="oper">
		<option value=""  <%if oper="" then response.write "selected"%>></option>
		<option value="I" <%if oper="I" then response.write "selected"%>>Inclus�o</option>
		<option value="A" <%if oper="A" then response.write "selected"%>>Altera��o</option>
		<option value="E" <%if oper="E" then response.write "selected"%>>Exclus�o</option>
		<option value="2" <%if oper="2" then response.write "selected"%>>2� Via</option>
		</select></td>
	<td class=titulo><select size="1" name="uoper">
		<option value=""  <%if uoper="" then response.write "selected"%>></option>
		<option value="I" <%if uoper="I" then response.write "selected"%>>Inclus�o</option>
		<option value="A" <%if uoper="A" then response.write "selected"%>>Altera��o</option>
		<option value="E" <%if uoper="E" then response.write "selected"%>>Exclus�o</option>
		<option value="2" <%if uoper="2" then response.write "selected"%>>2� Via</option>
		</select></td>
	<td class=titulo><input type="checkbox" name="compr" value="ON" <%=obs1 %>></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Altera��es  " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Altera��es" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
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
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lan�amento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lan�amento N�o pode ser alterado!');</script>"
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