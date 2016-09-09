<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a88")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Outros empregos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql = "INSERT INTO pfunc_empregos (" 
		sql = sql & "chapa, empresa, cargo, desde, ate, "
		sql = sql & "usuarioe, datae "
		sql = sql & ") "
		sql2 = " SELECT "
		sql2=sql2 & " '" & request.form("chapa") & "', "
		sql2=sql2 & " '" & request.form("empresa") & "', "
		sql2=sql2 & " '" & request.form("cargo") & "', "
		sql2=sql2 & " '" & request.form("desde") & "', "
		sql2=sql2 & " '" & request.form("ate") & "', "
		sql2=sql2 & " '" & session("usuariomaster") & "', "
		sql2=sql2 & "getdate()"
		sql1 = sql & sql2 & ""
		'response.write "<font size='2'>" & sql1
		conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
end if

if request.form="" then
%>
<form method="POST" action="empregos_nova.asp" name="planodesaude" >
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Outros empregos</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=fundo><select size="1" name="chapa" class=a>
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where chapa='" & request("chapa") & "'" 
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if request("chapa")=rs2("chapa") then tempc="selected" else tempc=""
%>
          <option value="<%=rs2("chapa")%>" <%=tempc%>><%=rs2("nome")%></option>
<%
rs2.movenext:loop:rs2.close
%>
	</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Empresa</td>
	<td class=titulo>Cargo</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="empresa" size="45" value=""></td>
	<td class=fundo><input type="text" name="cargo" size="30" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Desde:</td>
	<td class=titulo>Até:</td>
	<td class=titulo>&nbsp;</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="desde" size="10" value="">  </td>
	<td class=fundo><input type="text" name="ate" size="10" value="">  </td>
	<td class=fundo>&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
	<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
	<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
	<input type="button" value="Fechar" class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
else
'rs.close
end if
%>
<%
if request.form("bt_salvar")<>"" then
	Response.write "<p>Registro salvo.<br>"
	'response.write "<a href='javascript:window.close()'>Fechar Janela</a>"
%>
 <script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
<%
end if
%>
<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>