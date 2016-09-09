<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a99")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Menu</title>
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="javascript" type="text/javascript"><!--
function nomeu() {
	form.nome_autonomo.value=form.nome_autonomo.value.toUpperCase()
}
// --></script>
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

if request.form("bt_salvar")<>"" then
	tudook=1
	
	if request.form("idmenu")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Código: Informe o código do menu!');</script>"
	if request.form("menu")=""    then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Descrição: Informe um nome para o menu!');</script>"
	if request.form("sigla")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Sigla: Informe uma sigla para o menu!');</script>"

	sql = "INSERT INTO intranet_menus (idmenu, menu, sigla) "

	sql2 = " SELECT " & request.form("idmenu") & " "
	sql2=sql2 & ", '" & request.form("menu") & "' "
	sql2=sql2 & ", '" & request.form("sigla") & "' "
	sql1 = sql & sql2 & ""
	'response.write sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
else 'request.form=""
end if

'if request.form="" then
%>
<form method="POST" action="menus_nova.asp" name="form">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr><td class=grupo>Inclusão de Menu</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Código</td>
	<td class=titulo>Nome do Menu</td>
	<td class=titulo>Sigla</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="idmenu" size="5" value="<%=request.form("idmenu")%>"></td>
	<td class=fundo><input type="text" name="menu" size="35" value="<%=request.form("menu")%>"></td>
	<td class=fundo><input type="text" name="sigla" size="6" value="<%=request.form("sigla")%>" ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'"></td>
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
'end if
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