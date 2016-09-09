<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a68")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Candidato</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="javascript" type="text/javascript"><!--
function nomeu() {
	form.nome_candidato.value=form.nome_candidato.value.toUpperCase()
}
// --></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		sql = "INSERT INTO rs_candidato (id_requisicao, nome_candidato, idade, telefone, email "
		sql = sql & ") "
		sql2 = " SELECT " & request.form("id_requisicao") & ", '" & _
		request.form("nome_candidato") & "', " & request.form("idade") & ", '" & request.form("telefone") & "', '" & request.form("email") & "' "
		sql1 = sql & sql2 & ""
		'response.write "<font size='1'>" & sql1
		if tudook=1 then conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
	
end if

if request.form="" or (request.form<>"" and tudook=0) then
%>
<form method="POST" action="candidato_nova.asp" name="form">
<input type="hidden" name="id_requisicao" size="4" value="<%=request("codigo")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Candidato a vaga: <%=request("vaga")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Nome do Candidato</td>
</tr>
<tr>
	<td class=titulo>0</td>
	<td class=fundo><input type="text" name="nome_candidato" size="70" value="<%=request.form("nome_candidato")%>" onchange="nomeu()"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Idade</td>
	<td class=titulo>Telefone</td>
	<td class=titulo>Email</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="idade"  size="5" value="<%=request.form("idade")%>"></td>
	<td class=titulo><input type="text" name="telefone" size="15" value="<%=request.form("telefone")%>"></td>
	<td class=titulo><input type="text" name="email" size="40" value="<%=request.form("email")%>"></td>
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