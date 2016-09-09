<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Veículo</title>
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

if request.form("bt_salvar")<>"" then
	tudook=1
	sql="UPDATE veiculos_a SET "
	sql=sql & "obs    = '"   & request.form("obs")   & "', "
	if request.form("inicio")<>"" then 
		sql=sql & "inicio='" & dtaccess(request.form("inicio")) & "', "
	else
		sql=sql & "inicio=null, "
	end if
	if request.form("termino")<>"" then 
		sql=sql & "termino='" & dtaccess(request.form("termino")) & "', "
	else
		sql=sql & "termino=null, "
	end if
	sql=sql & "chapa   ='" & request.form("chapa") & "', "
	sql=sql & "cartao  ='" & request.form("cartao") & "', "
	sql=sql & "usuarioa='" & session("usuariomaster") & "', "
	if request.form("vy")="ON" then sql=sql & "vy = -1, " else sql=sql & "vy = 0, " 
	if request.form("bp")="ON" then sql=sql & "bp = -1, " else sql=sql & "bp = 0, " 
	if request.form("ns")="ON" then sql=sql & "ns = -1, " else sql=sql & "ns = 0, " 
	if request.form("jw")="ON" then sql=sql & "jw = -1, " else sql=sql & "jw = 0, " 
	if request.form("pavy")="ON" then sql=sql & "pavy = -1, " else sql=sql & "pavy = 0, " 
	if request.form("pabp")="ON" then sql=sql & "pabp = -1, " else sql=sql & "pabp = 0, " 
	if request.form("pans")="ON" then sql=sql & "pans = -1, " else sql=sql & "pans = 0, " 
	if request.form("pajw")="ON" then sql=sql & "pajw = -1, " else sql=sql & "pajw = 0, " 
	sql=sql & "dataa   =getdate() "
	sql=sql & " WHERE id_est=" & session("id_alt_est")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM veiculos_a WHERE id_est=" & session("id_alt_est")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_est=session("id_alt_est")
	else
		id_est=request("codigo")
	end if
	sqla="select * from veiculos_a where id_est=" & id_est
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
end if


if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_est")=rs("id_est")

sqlz="select nome from (select chapa, nome, codsecao as descricao, codsindicato from grades_novos union all select f.chapa collate database_default, f.nome collate database_default, f.secao collate database_default, f.codsindicato collate database_default from qry_funcionarios f) as t where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
%>
<form method="POST" action="estac_alteracao.asp" name="planodesaude">
<input type="hidden" name="id_est" size="4" value="<%=rs("id_est")%>" style="font-size: 8 pt">
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Veículo</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=titulo><p class=realce><%=rs("chapa")%> - <%=rsnome("nome")%></p></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Inicio</td>
	<td class=titulo>Término</td>
	<td class=titulo>Cartão B.P.</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="inicio" size="8" value="<%=rs("inicio")%>"></td>
	<td class=titulo><input type="text" name="termino" size="8" value="<%=rs("termino")%>"></td>  <!-- onfocus="this.blur()" -->
	<td class=titulo><input type="text" name="cartao" size="5" value="<%=rs("cartao")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Coral</td>
	<td class=titulo>Branco</td>
	<td class=titulo>Narc.</td>
	<td class=titulo>J.W.</td>
	<td class=titulo>Obs.</td>
<%if session("usuariomaster")="02379" then%>
	<td class=titulo>Ant.VY</td>
	<td class=titulo>Ant.BP</td>
	<td class=titulo>Ant.NS</td>
	<td class=titulo>Ant.JW</td>
<%end if%>
</tr>
<tr>
<%
emissao=formatdatetime(now,2)
if rs("vy")=0 then vy="" else vy="checked"
if rs("bp")=0 then bp="" else bp="checked"
if rs("ns")=0 then ns="" else ns="checked"
if rs("jw")=0 then jw="" else jw="checked"
if rs("pavy")=0 then pavy="" else pavy="checked"
if rs("pabp")=0 then pabp="" else pabp="checked"
if rs("pans")=0 then pans="" else pans="checked"
if rs("pajw")=0 then pajw="" else pajw="checked"
%>
	<td class=titulo><input type="checkbox" name="vy" value="ON" <%=vy%>></td>
	<td class=titulo><input type="checkbox" name="bp" value="ON" <%=bp%>></td>
	<td class=titulo><input type="checkbox" name="ns" value="ON" <%=ns%>></td>
	<td class=titulo><input type="checkbox" name="jw" value="ON" <%=jw%>></td>
	<td class=titulo><input type="text" name="obs" size="30" value="<%=rs("obs")%>"></td>
<%if session("usuariomaster")="02379" then%>
	<td class=titulo><input type="checkbox" name="pavy" value="ON" <%=pavy%>></td>
	<td class=titulo><input type="checkbox" name="pabp" value="ON" <%=pabp%>></td>
	<td class=titulo><input type="checkbox" name="pans" value="ON" <%=pans%>></td>
	<td class=titulo><input type="checkbox" name="pajw" value="ON" <%=pajw%>></td>
<%else%>
<%end if%>
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