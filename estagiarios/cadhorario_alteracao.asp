<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Horário - Estagiário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function renovacao1()	{ form.urenovacao.value=form.renovacao_anterior.value;	}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(4), varcur(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	if request.form("ativo")="ON" then ativo1 = 1 else ativo1 = 0
	jsemh=request.form("jsem_h")
	jsemm=request.form("jsem_m")
	jmesh=request.form("jmes_h")
	jmesm=request.form("jmes_m")
	if jsemh="" then jsemh=0
	if jsemm="" then jsemm=0
	if jmesh="" then jmesh=0
	if jmesm="" then jmesm=0
	jsem=(jsemh*60)+jsemm
	jmes=(jmesh*60)+jmesm
	sql="UPDATE est_cadhorario SET "
	sql=sql & "descricao = '" & request.form("descricao") & "', "
	if request.form("datacriacao")<>"" then 
		sql=sql & "datacriacao = '" & dtaccess(request.form("datacriacao")) & "', "
	else
		sql=sql & "datacriacao = null, "
	end if
	sql=sql & "jsem = " & jsem & ", "
	sql=sql & "jmes = " & jmes & ", "
	sql=sql & "ativo = "  & ativo1 & ", "
	sql=sql & "usuarioa='" & session("usuariomaster") & "', "
	sql=sql & "dataa   =getdate() "
	sql=sql & " WHERE codigo='" & session("id_alt_cadhor") & "' "
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM est_cadhorario WHERE codigo='" & session("id_alt_cadhor") & "' "
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null or request("codigo")="" then
		id_cadhor=session("id_alt_cadhor")
		if session("id_alt_cadhor")="" then id_cadhor=request.form("id_cadhor")
	else
		id_cadhor=request("codigo")
	end if
	sqla="select * from est_cadhorario "
	sqlb="where codigo='" & id_cadhor & "' "
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
session("id_alt_cadhor")=rs("codigo")
jsemh=int(rs("jsem")/60)
jsemm=rs("jsem")-jsemh*60
jmesh=int(rs("jmes")/60)
jmesm=rs("jmes")-jmesh*60
if request.form("codigo")=""      then codigo=rs("codigo")           else codigo=request.form("codigo")
if request.form("descricao")=""   then descricao=rs("descricao")     else descricao=request.form("descricao")
if request.form("datacriacao")="" then datacriacao=rs("datacriacao") else datacriacao=request.form("datacriacao")
if request.form("jsem_h")=""      then jsem_h=jsemh                  else jsem_h=request.form("jsem_h")
if request.form("jsem_m")=""      then jsem_m=jsemm                  else jsem_m=request.form("jsem_m")
if request.form("jmes_h")=""      then jmes_h=jmesh                  else jmes_h=request.form("jmes_h")
if request.form("jmes_m")=""      then jmes_m=jmesm                  else jmes_m=request.form("jmes_m")
if request.form("ativo")=""       then ativo=rs("ativo")             else ativo=request.form("ativo")
if ativo<>0 or ativo=true or ativo="ON" then ativo1="checked" else ativo1=""

%>
<form method="POST" action="cadhorario_alteracao.asp" name="form">
<input type="hidden" name="id_cadhor" size="4" value="<%=rs("codigo")%>" >  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Horário</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Código</td>
	<td class=titulo>Descrição</td>
</tr>
<tr>
	<td class=titulo><input type="hidden" name="codigo" size="9" value="<%=codigo%>" class="form_input"><%=codigo%></td>
	<td class=fundo><input type="text" name="descricao" size="50" value="<%=descricao%>" class="form_input"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Data Criação</td>
	<td class=titulo>Jorn.Semanal</td>
	<td class=titulo>Jorn.Mensal</td>
	<td class=titulo>Ativo</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="datacriacao" size="8" value="<%=datacriacao%>" class="form_input"></td>
	<td class=fundo><input type="text" name="jsem_h" size="2" value="<%=jsem_h%>" class="form_input"> <b>h
	<input type="text" name="jsem_m" size="2" value="<%=jsem_m%>" class="form_input"> m	
	</td>
	<td class=fundo><input type="text" name="jmes_h" size="3" value="<%=jmes_h%>" class="form_input"> <b>h
	<input type="text" name="jmes_m" size="2" value="<%=jmes_m%>" class="form_input"> m
	</td>
	<td class=fundo><input type="checkbox" name="ativo" value="ON" <%=ativo1%>>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo colspan=3>&nbsp;</td></tr>
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
end if
set rs=nothing
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