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
<title>Inclusão de Horário-Estagiário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function chapa1() {	form.chapa.value=form.nome.value;	}
function nome1() {	form.nome.value=form.chapa.value;	}
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

if request.form<>"" then
if request.form("bt_salvar")<>"" then
	tudook=1
	'response.write request.form
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
	sql = "INSERT INTO est_cadhorario (" 
	sql = sql & "codigo, descricao, datacriacao, jsem, jmes, ativo, usuarioc, datac "
	sql = sql & ") "
	sql = sql & " SELECT "
	sql = sql & " '" & request.form("codigo") & "', "
	sql = sql & " '" & request.form("descricao") & "', "
	if request.form("datacriacao")="" then sql =sql & "null," else sql =sql & " '" & dtaccess(request.form("datacriacao")) & "', "
	sql = sql & " '" & jsem & "', "
	sql = sql & " '" & jmes & "', "
	sql = sql  & " " & ativo1 & ", "
	sql = sql  & " '" & session("usuariomaster") & "', "
	sql = sql  & " getdate() "
	sql1 = sql
	'response.write "<font size='1'>" & sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
end if
else 'request.form=""
end if

'if request.form="" then
if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then

codigo     =request.form("codigo")
descricao  =request.form("descricao")
datacriacao=request.form("datacriacao")
jsem_h     =request.form("jsem_h")
jsem_m     =request.form("jsem_m")
jmes_h     =request.form("jmes_h")
jmes_m     =request.form("jmes_m")
ativo      =request.form("ativo")
if ativo="ON" then ativo1="checked" else ativo1=""

%>
<form method="POST" action="cadhorario_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Horário</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Código</td>
	<td class=titulo>Descrição</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="codigo" size="9" value="<%=codigo%>" class="form_input"></td>
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
	<td class=fundo><input type="hidden" name="datacriacao" size="8" value="<%=formatdatetime(now,2)%>"><%=formatdatetime(now,2)%></td>
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
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
else
'rs.close
set rs=nothing
set rsc=nothing
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