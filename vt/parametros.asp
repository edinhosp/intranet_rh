<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
if session("a70")="N" or session("a70")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Parâmetros de Configuração para Vale-Transporte</title>
<link rel="stylesheet" type="text/css" href="..\<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->
<%
'se for alterar registro por cookie
'response.cookies("vrh06")("registropagina")="25"

'para ler o cookie
rp=request.cookies("vrh06")("registropagina")

dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	if tudook=1 then
		'sql="update usuarios set password='" & request.form("novasenha") & "' where usuario='" & session("usuariomaster") & "' "
		'conexao.execute sql
		'response.write "<script language='JavaScript' type='text/javascript'>alert('Senha alterada!');</script>"
	end if
	for a=0 to request.form("ptotal")
		parametro=request.form("p" & a)
		novovalor=request.form("v" & a)
		if novovalor<>"" then
			novovalor=replace(novovalor,",",".")
			sql="update iParametros set valor=" & novovalor & " where parametro='" & parametro & "'"
			conexao.execute sql
		end if
	next
end if

sql="SELECT modulo, descricao, parametro, valor " & _
"FROM iParametros order by modulo, descricao"
rs.Open sql, ,adOpenStatic, adLockReadOnly

%>
<form method="POST" action="parametros.asp" name="form">
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse">
<tr><td class="grupo" colspan="4">Configurações</td></tr>
<tr>
	<td class="fundo">Módulo</td>
	<td class="fundo">Descrição</td>
	<td class="fundo">Valor Atual</td>
	<td class="fundo">Novo valor</td>
</tr>
<%
item=0
do while not rs.eof
%>
<tr>
	<td class="campo"><%=rs("modulo")%></td>
	<td class="campo"><%=rs("descricao")%></td>
<input type="hidden" name="p<%=item%>" value="<%=rs("parametro")%>">
	<td class="campo"><%=rs("valor")%></td>
	<td class="campo"><input type="text" name="v<%=item%>" size=6 value=""></td>
</tr>
<%
item=item+1
rs.movenext
loop
rs.close
%>
<input type="hidden" name="ptotal" value="<%=item-1%>">

</table>
<br>
<input type="submit" value="Salvar Alteração" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
</form>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>

<!-- -->
</td><td valign="top">

</td></tr></table>
<!-- -->
</body>
</html>