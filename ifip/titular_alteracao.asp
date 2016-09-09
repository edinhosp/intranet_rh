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
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Alteração de Titular</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>
</head>
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
		sql="UPDATE ifip_titulares SET "
		sql=sql & "tp_docente    ='" & request.form("tp_docente") & "', "
		sql=sql & "chapa         ='" & request.form("chapa")       & "' "
		sql=sql & "WHERE id_tit=" & session("id_alt_tit")
		conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		sql="DELETE FROM ifip_titulares WHERE id_tit=" & session("id_alt_tit")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
	if request("codigo")=null then
		id_tit=session("id_alt_tit")
	else
		id_tit=request("codigo")
	end if
	sqla="select * from ifip_titulares "
	sqlb="where id_tit=" & id_tit
	sql1=sqla & sqlb 
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if
%>

<%
if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_tit")=rs("id_tit")
%>
<form method="POST" action="titular_alteracao.asp" name="form">
<input type="hidden" name="id_tit" size="4" value="<%=rs("id_tit")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Alteração de Professor</td></tr>
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
if rs("tp_docente")=rs2("id_titular") then temptp="selected" else temptp=""
%>
	<option value="<%=rs2("id_titular")%>" <%=temptp%>><%=rs2("desc_titular")%></option>
<%
rs2.movenext:loop
rs2.close
%>
	</select>	  
	</td>
	<td class=fundo>
	<input type="text" name="chapa" size="5" value="<%=rs("chapa")%>" onchange="chapa1()">
	<select size="1" name="nome" onchange="nome1()">
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and codtipo='N' and codsindicato='03' order by nome"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if rs("chapa")=rs2("chapa") then tempch="selected" else tempch=""
%>
	<option value="<%=rs2("chapa")%>" <%=tempch%>><%=rs2("nome")%></option>
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
		<input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar"></td>
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

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	Response.write "<p>Registro atualizado."
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar janela" class="button" onClick="top.window.close()">
</form>
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>