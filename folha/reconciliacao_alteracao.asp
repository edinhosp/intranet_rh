<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Reconciliação</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1() { form.nome.value=form.nome.value.toUpperCase() }
--></script>
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

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		'if request.form("bolsa")="ON" then bolsa=-1 else bolsa=0
		sql="UPDATE reconciliacao SET " & _
		"id_tipo      = " & request.form("id_tipo") & " " & _
		",data        ='" & dtaccess(request.form("data")) & "' " & _
		",valor       = " & nraccess(request.form("valor")) & " " & _
		",anocomp     = " & request.form("anocomp") & " " & _
		",mescomp     = " & request.form("mescomp") & " " & _
		",obs         ='" & request.form("obs") & "' " & _
		" WHERE id_rec=" & session("id_alt_reco")
		'",usuarioa='" & session("usuariomaster") & "' "
		'",dataa   =getdate() "
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="UPDATE apont_adm set deletada=-1 WHERE id_adm=" & session("id_alt_adm")
		sql="DELETE FROM reconciliacao WHERE id_rec=" & session("id_alt_reco")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_reco=session("id_alt_reco")
		id_reco=request.form("id_reco")
	else
		id_reco=request("codigo")
	end if
	sql="select * from reconciliacao where id_rec=" & id_reco
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_reco")=rs("id_rec")
'response.write request.form
%>
<form method="POST" action="reconciliacao_alteracao.asp" name="form">
<input type="hidden" name="id_reco" size="4" value="<%=rs("id_rec")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
	<tr><td class=grupo>Alteração de Reconciliação <%=rs("id_rec")%></td></tr>
</table>

<!-- tipo lançamento -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Tipo de lançamento</td>
	<td class=titulo>Data</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="id_tipo" onfocus="javascript:window.status='Selecione o tipo de lançamento'">
<%
sqla="select id_tipo, tipo from reconciliacao_eventos order by tipo"
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if rs("id_tipo")=rsd("id_tipo") then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("id_tipo")%>" <%=tempc%>><%=rsd("tipo")%></option>
<%
rsd.movenext:loop
rsd.close
%>
	</select></td>
	<td class=fundo><input type="text" name="data" value="<%=rs("data")%>" size="9">
	</td>
</tr>
</table>

<!-- valor / competencia -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Valor R$</td>
	<td class=titulo>Mês/Ano Comp.</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="valor" value="<%=rs("valor")%>" size="12"></td>
	<td class=fundo><input type="text" name="mescomp" value="<%=rs("mescomp")%>" size="2">
	 / <input type="text" name="anocomp" value="<%=rs("anocomp")%>" size="4">
	</td>
</tr>
</table>

<!-- observacao -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Observação</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=rs("obs")%>" name="obs" size="70"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
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

conexao.close
set conexao=nothing
%>
</body>
</html>