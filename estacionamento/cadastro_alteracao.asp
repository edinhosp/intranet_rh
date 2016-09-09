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
<title>Alteração de Veículo</title>
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
	sql="UPDATE veiculos SET "
	sql=sql & "marca  = '"   & request.form("marca") & "', "
	sql=sql & "modelo = '"   & request.form("modelo")& "', "
	sql=sql & "ano    = '"   & request.form("ano")   & "', "
	sql=sql & "cor    = '"   & request.form("cor")   & "', "
	sql=sql & "placa  = '"   & request.form("placa") & "', "
	if request.form("dtcadastro")<>"" then 
		sql=sql & "dtcadastro='" & dtaccess(request.form("dtcadastro")) & "', "
	else
		sql=sql & "dtcadastro=null, "
	end if
	if request.form("dttermino")<>"" then 
		sql=sql & "dttermino='" & dtaccess(request.form("dttermino")) & "', "
	else
		sql=sql & "dttermino=null, "
	end if
	sql=sql & "chapa   ='" & request.form("chapa") & "', "
	sql=sql & "usuarioa='" & session("usuariomaster") & "', "
	sql=sql & "dataa   =getdate() "
	sql=sql & " WHERE id_veiculo=" & session("id_alt_veiculo")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM veiculos WHERE id_veiculo=" & session("id_alt_veiculo")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_veiculo=session("id_alt_veiculo")
	else
		id_veiculo=request("codigo")
	end if
	sqla="select * from veiculos "
	sqlb="where id_veiculo=" & id_veiculo
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if


if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_veiculo")=rs("id_veiculo")

sqlz="select nome from grades_aux_prof where chapa='" & rs("chapa") & "'"
sqlz="select chapa, nome from grades_aux_prof where chapa='" & rs("chapa") & "' " & _
"union all " & _
"select chapa collate database_default, nome collate database_default from qry_funcionarios where codsituacao<>'D' and chapa='" & rs("chapa") & "'"

set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
if rsnome.recordcount>0 then nome=rsnome("nome") else nome=""
%>
<form method="POST" action="cadastro_alteracao.asp" name="planodesaude">
<input type="hidden" name="id_veiculo" size="4" value="<%=rs("id_veiculo")%>" style="font-size: 8 pt">
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Veículo</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=titulo><p class=realce><%=rs("chapa")%> - <%=nome%></p></td></tr>
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
	rs2.movefirst
	do while not rs2.eof
	if rs2("marca")=rs("marca") then tempt="selected" else tempt=""
%>
		<option value="<%=rs2("marca")%>" <%=tempt%>><%=rs2("marca")%></option>
<%
	rs2.movenext
	loop
	rs2.close
%>
	</select></td>
	<td class=titulo><input type="text" name="modelo" size="25" value="<%=rs("modelo")%>"></td>
	<td class=titulo><input type="text" name="placa" size="12" value="<%=rs("placa")%>"></td>
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
	<td class=titulo><input type="text" name="cor"   size="12" value="<%=rs("cor")%>">  </td>
	<td class=titulo><input type="text" name="ano"   size="8"  value="<%=rs("ano")%>">  </td>
	<td class=titulo><input type="text" name="dtcadastro" size="8" value="<%=rs("dtcadastro")%>"></td>
	<td class=titulo><input type="text" name="dttermino"  size="8" value="<%=rs("dttermino")%>"> </td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
	<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
	<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
	<input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
	</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if
set rsc=nothing
set rsnome=nothing
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