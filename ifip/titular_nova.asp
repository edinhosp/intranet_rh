<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a30")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Titular</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql = "INSERT INTO ifip_titulares (id_ifip, tp_docente, chapa "
		sql = sql & ") "
		sql2 = " SELECT " & request.form("id_ifip") & ", '" & _
		request.form("tp_docente") & "', '" & _
		request.form("chapa") & "' "
		sql1 = sql & sql2 & ""
		'response.write "<font size='1'>" & sql1
		conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
	
end if

if request.form="" then
%>
<form method="POST" action="titular_nova.asp" name="form">
<input type="hidden" name="id_ifip" size="4" value="<%=request("codigo")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Professor</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Tipo de Participação</td>
	<td class=titulo>Chapa e Nome</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="tp_docente">
<%
sql2="select * from ifip_wtitular"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
		<option value="<%=rs2("id_titular")%>"><%=rs2("desc_titular")%></option>
<%
rs2.movenext:loop
rs2.close
%>
	</select>	  
	</td>
	<td class=fundo>
	<input type="text" name="chapa" size="5" onchange="chapa1()" value="">
	<select size="1" name="nome" onchange="nome1()">
	<option value=0>Selecione o professor</option>
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and codtipo='N' and codsindicato='03' order by nome"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("chapa")%>"><%=rs2("nome")%></option>
<%
rs2.movenext:loop
rs2.close
%>
	</select>	  
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
else
'rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" then
	Response.write "<p>Registro salvo.<br>"
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>