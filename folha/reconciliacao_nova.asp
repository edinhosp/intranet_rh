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
<title>Inclusão de Reconciliação</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {
	form.nome.value=form.nome.value.toUpperCase()
}
--></script>
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

if request.form("id_tipo")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione o tipo de lançamento!');</script>"
if request.form("valor")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o valor!');</script>"
'if request.form("bolsa")="ON" then bolsa=-1 else bolsa=0
		sqla = "INSERT INTO reconciliacao (id_tipo, data, valor, anocomp, mescomp, obs ) "
		
		sqlb = " SELECT " & request.form("id_tipo") & "" & _
		", '" & dtaccess(request.form("data")) & "'" & _
		", " & nraccess(request.form("valor")) & "" & _
		", " & request.form("anocomp") & "" & _
		", " & request.form("mescomp") & "" & _
		",'" & request.form("obs") & "' " 
		'sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		'sqlb=sqlb & ",getdate()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if 'request btsalvar
else 'request.form=""
end if

'if request.form("bt_salvar")="" then
%>
<form method="POST" action="reconciliacao_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
	<tr><td class=grupo>Inclusão de reconciliacao</td></tr>
</table>
<!-- tipo lancamento -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Tipo de lançamento</td>
	<td class=titulo>Data</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="id_tipo" onfocus="javascript:window.status='Selecione o tipo de evento'">
<%
sqla="select id_tipo, tipo from reconciliacao_eventos order by tipo"
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if request.form("id_tipo")=rsd("id_tipo") then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("id_tipo")%>" <%=tempc%>><%=rsd("tipo")%></option>
<%
rsd.movenext:loop
rsd.close

if isdate(request.form("data"))=true then data1=formatdatetime(request.form("data"),2) else data1=request.form("data")
%>
	</select></td>
	<td class=fundo><input type="text" name="data" value="<%=data1%>" size="9">
	</td>
</tr>
</table>

<!-- valor / competência -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Valor R$</td>
	<td class=titulo>Mês/Ano Comp.</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="valor" value="<%=request.form("valor")%>" size="12"></td>
	<td class=fundo><input type="text" name="mescomp" value="<%=request.form("mescomp")%>" size="2">
	 / <input type="text" name="anocomp" value="<%=request.form("anocomp")%>" size="4">
	</td>
</tr>
</table>

<!-- observacao -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Observação</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=request.form("obs")%>" name="obs" size="70"></td>
</tr>
</table>


<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
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
'end if   'request.form=""
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if
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