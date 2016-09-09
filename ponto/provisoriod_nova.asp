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
<title>Inclusão de Devolução de Provisório</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
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

if request.form("provisorio")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione o número do provisorio!');</script>"
if request.form("chapa")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione um funcionário!');</script>"
'if request.form("bolsa")="ON" then bolsa=-1 else bolsa=0
		hora=hour(request.form("horae"))*60
		minuto=minute(request.form("horae"))
		horae=hora+minuto
		sqla = "INSERT INTO provisorio (operacao, provisorio, chapa, datae, horae, usuarioc, datac ) "
		
		sqlb = " SELECT 'D','" & request.form("provisorio") & "'" & _
		", '" & request.form("chapa") & "'" & _
		", '" & dtaccess(request.form("datae")) & "'" & _
		", " & horae & "" & _
		",'" & session("usuariomaster") & "'" & _
		",getdate()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if 'request btsalvar
else 'request.form=""
end if

'if request.form("bt_salvar")="" then
%>
<form method="POST" action="provisoriod_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
	<tr><td class=grupo>Inclusão de Devolução de Provisório</td></tr>
</table>
<!-- tipo lancamento -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
<tr>
	<td class=titulo>Código Provisório</td>
	<td class=titulo>Entregue a</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="provisorio" onfocus="javascript:window.status='Selecione o tipo de evento'">
<%
sqla="select a.codcracha from acracha a where a.situacao=1 and a.codcracha not in ( " & _
"select u.codcracha from ausoprov u where now() between u.datainicio and u.datafim ) "
sqla="select u.chapafunc, u.codcracha, u.datainicio, u.datafim from corporerm.dbo.ausoprov u where getdate() between u.datainicio and u.datafim " & _
"order by u.codcracha, u.datainicio "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione....</option>"
if rsd.recordcount>0 then
rsd.movefirst:do while not rsd.eof
if request("provisorio")=rsd("codcracha") then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("codcracha")%>" <%=tempc%>><%=rsd("codcracha")%></option>
<%
rsd.movenext:loop
end if
rsd.close
%>
	</select></td>
	<td class=fundo><input type="text" value="<%=request("chapa")%>" name="chapa" size="5" onfocus="javascript:window.status='Informe o chapa do funcionário'" onchange="chapa1()">
		<select size="1" name="nome" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" onchange="nome1()">
<%
sql2="select chapa, nome from (select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' union all select chapa collate database_default, nome collate database_default from provisorio_funcnovo) as t order by nome "
'if session("dp_chapa")<>"" then sql2=sql2 & "and chapa='" & session("dp_chapa") & "'" else sql2=sql2 & "order by nome"
rsd.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o funcionário....</option>"
rsd.movefirst:do while not rsd.eof
if request("chapa")=rsd("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rsd("chapa")%>" <%=temp%>><%=rsd("nome")%></option>
<%
rsd.movenext:loop
rsd.close
%>
		</select></td>
</tr>
</table>

<!-- valor / competência -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
<tr>
	<td class=titulo>Data Entrega</td>
	<td class=titulo>Hora Entrega</td>
	<td class=titulo>Entregue por</td>
</tr>
<tr>
	<td class=fundo height=25><input type="hidden" name="datae" value="<%=formatdatetime(now(),2)%>"><font color=blue><%=formatdatetime(now(),2)%></td>
	<td class=fundo><input type="hidden" name="horae" value="<%=formatdatetime(now(),4)%>"><font color=blue><%=formatdatetime(now(),4)%></td>
	<td class=fundo><input type="hidden" name="usuarioc" value="<%=session("usuariomaster")%>"><font color=blue><%=session("usuariomaster")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
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