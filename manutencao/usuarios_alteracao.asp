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
<title>Alteração de Usuário</title>
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
	if request.form("usuario")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Chapa: Informe a chapa do usuário!');</script>"
	if request.form("nome")=""    then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Nome : Informe o nome do usuário!');</script>"
	if request.form("password")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Senha: Informe a senha do usuário');</script>"
	if request.form("ativo")="ON" then valueativo=1 else valueativo=0
	if request.form("master")="ON" then valuemaster=1 else valuemaster=0
	if request.form("timeout")="" then valuetimeout=45 else valuetimeout=request.form("timeout")

	sql="UPDATE usuarios SET "
	sql=sql & " nome    = '" & request.form("nome") & "' "
	sql=sql & ",password= '" & request.form("password") & "' "
	sql=sql & ",grupo   = '" & request.form("grupo") & "' "
	sql=sql & ",estilo  = '" & request.form("estilo") & "' "
	sql=sql & ",timeout = " & valuetimeout & " "
	sql=sql & ",ativo   = " & valueativo & " "
	sql=sql & ",master  = " & valuemaster & " "
	sql=sql & "WHERE usuario='" & session("id_alt_usuario") & "' "
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM usuarios WHERE usuario='" & session("id_alt_usuario") & "' "
		'response.write "<br><br><br><br><br><br><br><br><br><br><br>" & sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if isnull(request("codigo")) or request("codigo")="" then
		id_usuario=session("id_alt_usuario")
	else
		id_usuario=request("codigo")
	end if
	sql1="select * from usuarios where usuario='" & id_usuario & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
session("id_alt_usuario")=rs("usuario")
if rs("ativo")=true then valueativo="checked" else valueativo=""
if rs("master")=true then valuemaster="checked" else valuemaster=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
%>
<form method="POST" action="usuarios_alteracao.asp" name="form">
<input type="hidden" name="id_usuario" size="5" value="<%=rs("usuario")%>">  
<table border="0" cellpadding="3" cellspacing="0" width="500">
	<tr><td class=grupo>Alteração de Usuário</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Código/Chapa</td>
	<td class=titulo>Nome do Usuário</td>
	<td class=titulo>Senha</td>
	<td class=titulo>Ativo</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="usuario" size="5" value="<%=rs("usuario")%>"></td>
	<td class=fundo><input type="text" name="nome" size="25" value="<%=rs("nome")%>"></td>
	<td class=fundo><input type="password" name="password" size="6" alt="<%=rs("password")%>" value="<%=rs("password")%>" onfocus="javascript:window.status='<%=rs("password")%>'" ></td>
	<td class=fundo><input type="checkbox" name="ativo" value="ON" <%=valueativo%> ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Grupo</td>
	<td class=titulo>Estilo CSS</td>
	<td class=titulo>Timeout</td>
	<td class=titulo>Master</td>
</tr>
<tr>
	<td class=fundo><select name="grupo" size="1">
		<option value="">Selecione um grupo</option>
<%
sql2="select distinct grupo from usuarios order by grupo"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
	if rs("grupo")=rs2("grupo") then valuegrupo="selected" else valuegrupo=""
	response.write "<option " & valuegrupo & " value=""" & rs2("grupo") & """>" & rs2("grupo") & "</option>"
rs2.movenext
loop
rs2.close
%>
	</select>
	</td>
	<td class=fundo><input type="text" name="estilo" size="15" value="<%=rs("estilo")%>" onfocus="javascript:window.status=''"></td>
	<td class=fundo><input type="number" name="timeout" size="5" min="15" max="180" value="<%=rs("timeout")%>" ></td>
	<td class=fundo><input type="checkbox" name="master" value="ON" <%=valuemaster%> ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
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

conexao.close
set conexao=nothing
%>
</body>
</html>