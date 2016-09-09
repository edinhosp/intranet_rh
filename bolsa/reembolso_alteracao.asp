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
<title>Alteração de Reembolso</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function mensalidade1()	{
	mens=form.mensalidade.value;mens=mens.replace('.','');
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
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	sql="UPDATE bolsistas_reembolso SET "
	sql=sql & "mes_base       ='" & dtaccess(request.form("mes_base")) & "' "
	sql=sql & ",mensalidade   = " & nraccess(request.form("mensalidade")) & " "
	sql=sql & ",porcentagem   = " & nraccess(request.form("porcentagem")) & " "
	sql=sql & ",reembolso     = " & nraccess(request.form("reembolso")) & " "
	sql=sql & ",data_pagamento='" & dtaccess(request.form("data_pagamento")) & "' "
	'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
	'sql=sql & ",dataa   =getdate() "
	sql=sql & " WHERE id_reembolso=" & session("id_alt_reembolso")
	'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
	'response.write sql
	if tudook=1 then conexao.Execute sql, , adCmdText else tudook=0
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM bolsistas_reembolso WHERE id_reembolso=" & session("id_alt_reembolso")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_reembolso=session("id_alt_reembolso")
		id_reembolso=request.form("id_reembolso")
	else
		id_reembolso=request("codigo")
	end if
	sql="select * from bolsistas_reembolso where id_reembolso=" & id_reembolso
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_reembolso")=rs("id_reembolso")
sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
'response.write request.form
%>
<form method="POST" action="reembolso_alteracao.asp" name="form">
<input type="hidden" name="id_reembolso" size="4" value="<%=rs("id_reembolso")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Alteração de Reembolso de Mensalidade <%=rs("id_reembolso")%></td></tr>
</table>

<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
	<td class=titulo><%=rs("id_reembolso")%></td>
	<td class=titulo><%=rs("chapa")%></td>
	<td class=titulo><%=rsnome("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Mês base</td>
	<td class=titulo>Vr.Mensalidade</td>
	<td class=titulo>Perc.%</td>
</tr>
<tr>
<%mesbase=numzero(month(rs("mes_base")),2) & "/" & year(rs("mes_base"))%>
	<td class=fundo><input type="text" name="mes_base" size="9" value="<%=mesbase%>" ></td>
	<td class=fundo><input type="text" name="mensalidade" class=vr size="12" value="<%=formatnumber(rs("mensalidade"),2)%>" onchange="mensalidade1()" ></td>
	<td class=fundo><input type="text" name="porcentagem" size="5" value="<%=rs("porcentagem")%>" onchange="mensalidade1()" ></td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Reembolso</td>
	<td class=titulo>Data Pagamento</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="reembolso"      size="12" class=vr value="<%=formatnumber(rs("reembolso"),2)%>"></td>
	<td class=fundo><input type="text" name="data_pagamento" size="9" value="<%=rs("data_pagamento")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

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