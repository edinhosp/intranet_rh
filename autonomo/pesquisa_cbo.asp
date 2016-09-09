<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a52")="N" or session("a52")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pesquisa CBO-2002</title>
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
if request.form<>"" then
	dim conexao, conexao2, chapach, rs, rs2
	set conexao=server.createobject ("ADODB.Connection")
	conexao.Open Application("conexao")

	sqla="SELECT cbo, nome_ocupacao FROM temp_cbo "
	temp=request.form("loccbo")
	if isnumeric(temp) then
		sqlb="WHERE cbo like '%" & temp & "%'"
	else
		sqlb="WHERE nome_ocupacao like '%" & temp & "%'"
	end if
	sqlc="ORDER BY nome_ocupacao"
	sql1=sqla & sqlb & sqlc
	set rs=server.createobject ("ADODB.Recordset")
	Set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then session("loccbo")=rs("cbo") else session("loccbo")=temp
	end if
%>
<form method="POST" action="pesquisa_cbo.asp">
<table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" width="380">
<tr><td class=grupo colspan=2>Pesquisa de CBO-2002</td></tr>
<tr><td><font size=1>Localizar <input type="text" name="loccbo" size="15" class="form_box" value="<%=session("loccbo")%>">
	<input type="submit" value="Pesquisar" name="B1" class="button"></td>
	<td><input type="button" value="Fechar" class="button" onClick="top.window.close()">
	</td></tr></table>
</form>
<%
if request.form<>"" then
	if rs.recordcount>0 then
		session("loccbo")=rs("cbo")
%>
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="380">
<tr>
	<td class=titulo valign="middle">CBO</td>
	<td class=titulo valign="middle">Descrição</td>
</tr>
<%
		rs.movefirst
		do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("cbo")%></td>
	<td class=campo><%=rs("nome_ocupacao")%></td>
</tr>
<%
		rs.movenext
		loop
%>
</table>
<%
	else 'sem registros
%>
<p>
<b><font color="#FF0000">
Esta seleção não mostra nenhum registro.</font></b></p>
<%
	end if 'rs.recordcount
%>
</body>
</html>
<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
end if ' request.form
%>