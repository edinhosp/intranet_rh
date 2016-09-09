<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Veículo</title>
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

if request.form("bt_salvar")<>"" then
	tudook=1
	
	sql = "INSERT INTO veiculos (" 
	sql = sql & "chapa, marca, modelo, ano, cor, placa, "
	sql = sql & "dtcadastro, dttermino, usuarioa, dataa "
	sql = sql & ") "
	sql2 = " SELECT "
	sql2=sql2 & " '" & request.form("chapa") & "', "
	sql2=sql2 & " '" & request.form("marca") & "', "
	sql2=sql2 & " '" & request.form("modelo") & "', "
	sql2=sql2 & " '" & request.form("ano") & "', "
	sql2=sql2 & " '" & request.form("cor") & "', "
	sql2=sql2 & " '" & request.form("placa") & "', "
	if request.form("dtcadastro")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dtcadastro")) & "', "
	if request.form("dttermino")=""  then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dttermino")) & "', "
	sql2=sql2 & " '" & session("usuariomaster") & "', "
	sql2=sql2 & " getdate() "
	sql1 = sql & sql2 & ""
	'response.write "<font size='2'>" & sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
end if

%>
<form method="POST" action="cadastro_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Veículo</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=titulo><select size="1" name="chapa" class=a>
<%
sql2="select chapa, nome from grades_aux_prof where chapa='" & request("chapa") & "' " & _
"union " & _
"select chapa collate database_default, nome collate database_default from qry_funcionarios where codsituacao<>'D' and chapa='" & request("chapa") & "' " & _
"union " & _
"select chapa, nome from grades_novos where codsituacao<>'D' and codsindicato<>'03' "

rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if request("chapa")=rs2("chapa") then tempc="selected" else tempc=""
%>
          <option value="<%=rs2("chapa")%>" <%=tempc%>><%=rs2("nome")%></option>
<%
rs2.movenext
loop
rs2.close
%>
	</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Marca</td>
	<td class=titulo>Modelo</td>
	<td class=titulo>Placa</td>
</tr>
<tr>
	<td class=titulo><select size="1" name="marca">
		<option value="">Selecione a marca</option>
<%
	sqla="SELECT * from veiculos_marca "
	rs2.Open sqla, ,adOpenStatic, adLockReadOnly
	rs2.movefirst:do while not rs2.eof
%>
		<option value="<%=rs2("marca")%>"><%=rs2("marca")%></option>
<%
	rs2.movenext:loop
	rs2.close
%>
	</select></td>
	<td class=titulo><input type="text" name="modelo" size="25" value=""></td>
	<td class=titulo><input type="text" name="placa" size="12" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cor</td>
	<td class=titulo>Ano</td>
	<td class=titulo>Dt.Cadastro</td>
	<td class=titulo>Dt.Término</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="cor"   size="12" value="">  </td>
	<td class=titulo><input type="text" name="ano"   size="8"  value="">  </td>
	<td class=titulo><input type="text" name="dtcadastro" size="8" value="<%=formatdatetime(now,2)%>"></td>
	<td class=titulo><input type="text" name="dttermino"  size="8" value=""> </td>
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
'rs.close
set rs=nothing
set rs2=nothing
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