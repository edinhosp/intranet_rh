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
<title>Inclusão de Relatório</title>
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
		sql = "INSERT INTO ifip_relatorios (id_ifip, sequencia, periodicidade, dt_prevista, dt_apresentacao "
		sql = sql & ") "
		sql2 = " SELECT " & request.form("id_ifip") & ", '" & _
		request.form("sequencia") & "', '" & _
		request.form("periodicidade") & "', "
		if request.form("dt_prevista")=""     then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dt_prevista")) & "', "
		if request.form("dt_apresentacao")="" then sql2=sql2 & "null " else sql2=sql2 & " '" & dtaccess(request.form("dt_apresentacao")) & "' "
		sql1 = sql & sql2 & ""
		'response.write "<font size='1'>" & sql1
		conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
	
end if

if request.form="" then
%>
<form method="POST" action="relatorio_nova.asp" name="form">
<input type="hidden" name="id_ifip" size="4" value="<%=request("codigo")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Relatório</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Sequência</td>
	<td class=titulo>Periodicidade</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="sequencia" size="3" value="<%=request("sequencia")%>"></td>
	<td class=fundo>
		<select size="1" name="periodicidade">
<%
sql2="select * from ifip_wperiodicidade"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
		<option value="<%=rs2("id_periodicidade")%>"><%=rs2("desc_periodicidade")%></option>
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
	<td class=fundo><input type="text" name="dt_prevista" size="8" value=""></td>
	<td class=fundo><input type="text" name="dt_apresentacao" size="8" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
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
end if

if request.form("bt_salvar")<>"" then
	Response.write "<p>Registro salvo.<br>"
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>