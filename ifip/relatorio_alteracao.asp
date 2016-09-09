<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a30")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Alteração de Relatório</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql="UPDATE ifip_relatorios SET "
		sql=sql & "sequencia    ='" & request.form("sequencia")     & "', "
		sql=sql & "periodicidade='" & request.form("periodicidade") & "', "
		if request.form("dt_prevista")=""     then
			sql=sql & "dt_prevista=null,"
		else
			sql=sql & "dt_prevista='" & dtaccess(request.form("dt_prevista")) & "', "
		end if
		if request.form("dt_apresentacao")=""     then
			sql=sql & "dt_apresentacao=null "
		else
			sql=sql & "dt_apresentacao='" & dtaccess(request.form("dt_apresentacao")) & "' "
		end if
		sql=sql & "WHERE id_rel=" & session("id_alt_rel")
		conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		sql="DELETE FROM ifip_relatorios WHERE id_rel=" & session("id_alt_rel")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
	if request("codigo")=null then
		id_rel=session("id_alt_rel")
	else
		id_rel=request("codigo")
	end if
	sqla="select * from ifip_relatorios "
	sqlb="where id_rel=" & id_rel
	sql1=sqla & sqlb 
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_rel")=rs("id_rel")
%>
<form method="POST" action="relatorio_alteracao.asp" name="form">
<input type="hidden" name="id_rel" size="4" value="<%=rs("id_rel")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Alteração de Relatório</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Sequência</td>
	<td class=titulo>Periodicidade</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="sequencia" size="3" value="<%=rs("sequencia")%>"></td>
	<td class=fundo>
		<select size="1" name="periodicidade">
<%
sql2="select * from ifip_wperiodicidade"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if rs("periodicidade")=rs2("id_periodicidade") then temptp="selected" else temptp=""
%>
		<option value="<%=rs2("id_periodicidade")%>" <%=temptp%>><%=rs2("desc_periodicidade")%></option>
<%
rs2.movenext:loop
rs2.close
%>
		</select>	  
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Data prevista p/entrega</td>
	<td class=titulo>Data efetiva de entrega</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="dt_prevista" size="8" value="<%=rs("dt_prevista")%>"></td>
	<td class=fundo><input type="text" name="dt_apresentacao" size="8" value="<%=rs("dt_apresentacao")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
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
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	Response.write "<p>Registro atualizado."
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar janela" class="button" onClick="top.window.close()">
</form>
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>