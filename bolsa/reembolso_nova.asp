<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a65")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Reembolso</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function mensalidade1()	{
	mens=form.mensalidade.value;
	porc=form.porcentagem.value;
	mens=mens.replace(',','.')
	r=Math.ceil(parseFloat(mens)*parseFloat(porc)+0)/100
	r2=r.toString()
	r2=r2.replace('.',',')
	form.reembolso.value=r2
}
--></script>
<script language="VBScript">
	Sub mensalidade_onChange
		mens=document.form.mensalidade.value
		porc=document.form.porcentagem.value
		r=int(cdbl(mens)*cdbl(porc)+0.5)/100
		document.form.reembolso.value=r
	End Sub
	Sub porcentagem_onChange
		mens=document.form.mensalidade.value
		porc=document.form.porcentagem.value
		r=int(cdbl(mens)*cdbl(porc)+0.5)/100
		document.form.reembolso.value=r
	End Sub
</script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, ok
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		sqla = "INSERT INTO bolsistas_reembolso (chapa, id_bolsa, mes_base, mensalidade, porcentagem, reembolso, data_pagamento "
		sqla = sqla & " )"
		sqlb = " SELECT '" & request.form("chapa") & "'"
		sqlb=sqlb & "," & request.form("id_bolsa") & " "
		sqlb=sqlb & ",'" & dtaccess(request.form("mes_base")) & "' "
		sqlb=sqlb & "," & nraccess(request.form("mensalidade")) & " "
		sqlb=sqlb & "," & nraccess(request.form("porcentagem")) & " "
		sqlb=sqlb & "," & nraccess(request.form("reembolso")) & " "
		sqlb=sqlb & ",'" & dtaccess(request.form("data_pagamento")) & "' "
		'sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		'sqlb=sqlb & ",now()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText else tudook=0
	end if 'request btsalvar
end if

if request.form("bt_salvar")="" then
%>
<form method="POST" action="reembolso_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Inclusão de Reembolso Mensalidade</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
chapa=request("codigo")
id_bolsa=request("id")
porc=request("porc")
%>
<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=titulo>0</td>
	<td class=titulo><%=chapa%><input type="hidden" value="<%=chapa%>" name="chapa">
		<input type="hidden" value="<%=id_bolsa%>" name="id_bolsa"></td>
	<td class=titulo>
<%
sql2="select nome from corporerm.dbo.pfunc where chapa='" & chapa & "' "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
nome=rsc("nome")
rsc.close
%>
	<%=nome%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Mês base</td>
	<td class=titulo>Vr.Mensalidade</td>
	<td class=titulo>Perc.%</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="mes_base"    size="9" value=""></td>
	<td class=fundo><input type="text" name="mensalidade" size="12" class=vr value="0" onchange="mensalidade1()"></td>
	<td class=fundo><input type="text" name="porcentagem" size="5" value="<%=porc%>" onchange="mensalidade1()"></td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Reembolso</td>
	<td class=titulo>Data Pagamento</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="reembolso"      size="12" class=vr value="0"></td>
	<td class=fundo><input type="text" name="data_pagamento" size="9" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
	</td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2" onfocus="javascript:window.status='Clique para desfazer e limpar a tela'"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()" onfocus="javascript:window.status='Clique aqui para fechar sem salvar'"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing
end if   'request.form=""
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