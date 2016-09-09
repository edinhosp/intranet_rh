<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a81")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Dependente</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="javascript" type="text/javascript"><!--
function nome1() { form.dependente.value=form.dependente.value.toUpperCase()}
function nome2() { form.mae.value=form.mae.value.toUpperCase()}
// --></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(18)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	sql="UPDATE assmed_dep SET "
	sql=sql & "dependente    = '"   & request.form("dependente") & "', "
	sql=sql & "sexo          = '"   & request.form("sexo")       & "', "
	if request.form("nascimento")<>"" then 
		sql=sql & "nascimento = '"   & dtaccess(request.form("nascimento"))  & "', "
	else
		sql=sql & "nascimento = null, "
	end if
	sql=sql & "parentesco    = '"   & request.form("parentesco")    & "', "
	sql=sql & "cpf           = '"   & request.form("cpf")    & "', "
	sql=sql & "mae           = '"   & request.form("mae")      & "', "
	if request.form("dt_evento")<>"" then 
		sql=sql & "dt_evento = '"   & dtaccess(request.form("dt_evento")) & "' "
	else
		sql=sql & "dt_evento = null "
	end if
	sql=sql & " WHERE id_dep=" & session("id_alt_dep")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM assmed_dep WHERE id_dep=" & session("id_alt_dep")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null then
		id_dep=session("id_alt_dep")
	else
		id_dep=request("codigo")
	end if
	sqla="select * from assmed_dep "
	sqlb="where id_dep=" & id_dep
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_dep")=rs("id_dep")

sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
%>
<form method="POST" action="ctr_dep_alteracao.asp" name="form">
<input type="hidden" name="id_dep" size="4" value="<%=rs("id_dep")%>">  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Dependente de Assistência Médica</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário Titular</td></tr>
<tr><td class=titulo><p class=realce><%=rs("chapa")%> - <%=rsnome("nome")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Nome do dependente</td>
	<td class=titulo>Sexo      </td>
	<td class=titulo>Nascimento</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="dependente" size="45" value="<%=rs("dependente")%>" onchange="nome1()"></td>
	<td class=fundo><select size="1" name="sexo">
		<option selected>Selecione o sexo</option>
		<option value="F" <%if rs("sexo")="F" then response.write "selected"%>>Feminino</option>
		<option value="M" <%if rs("sexo")="M" then response.write "selected"%>>Masculino</option>
		</select></td>
	<td class=fundo><input type="text" name="nascimento" size="12" value="<%=rs("nascimento")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Parentesco</td>
	<td class=titulo>Nome da mãe</td>
	<td class=titulo>Data do Evento</td>
	<td class=titulo>CPF</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="parentesco">
<%
varpar(0)="Filha"
varpar(1)="Filho"
varpar(2)="Esposa"
varpar(3)="Esposo"
varpar(4)="Companheira"
varpar(5)="Companheiro"
varpar(6)="Agregado/a"
varpar(7)="Tutelado/a"
varpar(8)="Designado/a"
varpar(0)="Filha"
varpar(1)="Filho"
varpar(2)="Esposa"
varpar(3)="Esposo"
varpar(4)="Companheira"
varpar(5)="Companheiro"
varpar(6)="Agregado/a"
varpar(7)="Tutelado/a"
varpar(8)="Designado/a"
for a=1 to 10
	varpar(8+a)="Filho(a) " & a
next
for a=0 to 18
	if rs("parentesco")=varpar(a) then tempp="selected" else tempp=""
%>
	<option value="<%=varpar(a)%>" <%=tempp%>><%=varpar(a)%></option>
<%
next
%>
	</select></td>
	<td class=titulo><input type="text" name="mae" size="35" value="<%=rs("mae")%>" onchange="nome2()"></td>
	<td class=titulo><input type="text" name="dt_evento"  size="12" value="<%=rs("dt_evento")%>"></td>
	<td class=titulo><input type="text" name="cpf"  size="11" value="<%=rs("cpf")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
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
<%
%>