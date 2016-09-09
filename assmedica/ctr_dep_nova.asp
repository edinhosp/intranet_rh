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
<title>Inclusão de Dependente</title>
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
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	sql = "INSERT INTO assmed_dep (" 
	sql = sql & "chapa, dependente, sexo, nascimento, dt_evento, "
	sql = sql & "parentesco, cpf, mae "
	sql = sql & ") "
	sql2 = " SELECT "
	sql2=sql2 & " '" & request.form("chapa") & "', "
	sql2=sql2 & " '" & request.form("dependente") & "', "
	sql2=sql2 & " '" & request.form("sexo") & "', "
	if request.form("nascimento")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("nascimento")) & "', "
	if request.form("dt_evento")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dt_evento")) & "', "
	sql2=sql2 & " '" & request.form("parentesco") & "', "
	sql2=sql2 & " '" & request.form("cpf") & "', "
	sql2=sql2 & " '" & request.form("mae") & "' "
	sql1 = sql & sql2 & ""
	'response.write "<font size='2'>" & sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
end if

if request.form("bt_salvar")="" or (request.form("bt_salvar")<>"" and tudook=0) then
%>
<form method="POST" action="ctr_dep_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Dependente</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário Titular</td></tr>
<tr><td class=titulo><select size="1" name="chapa" class=a>
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' "
if request("chapa")<>"" then sql2=sql2 & "and chapa='" & request("chapa") & "'" else sql2=sql2 & "order by nome"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if request("chapa")=rsc("chapa") then tempc="selected" else tempc=""
%>
          <option value="<%=rsc("chapa")%>" <%=tempc%>><%=rsc("nome")%></option>
<%
rsc.movenext
loop
rsc.close
%>
	</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Nome do dependente</td>
	<td class=titulo>Sexo      </td>
	<td class=titulo>Nascimento</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="dependente" size="45" value="" onchange="nome1()"></td>
	<td class=titulo><select size="1" name="sexo">
		<option value="F">Feminino</option>
		<option value="M">Masculino</option>
		</select></td>
	<td class=titulo><input type="text" name="nascimento" size="12"></td>
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
	<td class=titulo><select size="1" name="parentesco">
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
for a=1 to 10
	varpar(8+a)="Filho(a) " & a
next
for a=0 to 18
%>
	<option value="<%=varpar(a)%>" <%=tempp%>><%=varpar(a)%></option>
<%
next
%>
	</select></td>
	<td class=titulo><input type="text" name="mae" size="35" onchange="nome2()"></td>
	<td class=titulo><input type="text" name="dt_evento"  size="12"></td>
	<td class=titulo><input type="text" name="cpf"  size="11"></td>
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