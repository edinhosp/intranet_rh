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
<title>Altera��o de Menu</title>
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
	tudook=1	
	if request.form("idmenu")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('C�digo: Informe o c�digo do menu!');</script>"
	if request.form("menu")=""    then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Descri��o: Informe um nome para o menu!');</script>"
	if request.form("sigla")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Sigla: Informe uma sigla para o menu!');</script>"

	sql="UPDATE intranet_menus SET "
	sql=sql & " idmenu   = 	" & request.form("idmenu") & " "
	sql=sql & ",menu     = '" & request.form("menu") & "' "
	sql=sql & ",sigla    = '" & request.form("sigla") & "' "
	sql=sql & "WHERE idmenu=" & session("id_alt_menu") & " "
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM intranet_menus WHERE idmenu=" & session("id_alt_menu")
		'response.write "<br><br><br><br><br><br><br><br><br><br><br>" & sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if isnull(request("idmenu")) or request("idmenu")="" then
		idmenu=session("id_alt_menu")
	else
		idmenu=request("idmenu")
	end if
	sql1="select * from intranet_menus where idmenu=" & idmenu & " "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
session("id_alt_menu")=rs("idmenu")
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
%>
<form method="POST" action="menus_alteracao.asp" name="form">
<input type="hidden" name="idmenu" size="5" value="<%=rs("idmenu")%>">  
<table border="0" cellpadding="3" cellspacing="0" width="500">
	<tr><td class=grupo>Altera��o de Menu</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>C�digo</td>
	<td class=titulo>Descri��o</td>
	<td class=titulo>Sigla</td>
</tr>
<tr>
	<td class=fundo><%=rs("idmenu")%></td>
	<td class=fundo><input type="text" name="menu" size="35" value="<%=rs("menu")%>"></td>
	<td class=fundo><input type="text" name="sigla" size="6" value="<%=rs("sigla")%>" ></td>
</tr>
</table>


<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Altera��es" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lan�amento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lan�amento N�o pode ser alterado!');</script>"
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

conexao.close
set conexao=nothing
%>
</body>
</html>
