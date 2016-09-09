<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a98")="N" or session("a98")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Arvore de Disciplinas</title>
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
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

areacon=request.form("areacon")
if areacon="" or isnull(areacon) then areacon=session("areacon")
session("areacon")=areacon
%>
<p class=realce>Relação de Disciplinas e Cursos</p>
<form action="treed.asp" method="post" name="form">
<p>Área de Conhecimento: 
<select name="areacon" onchange="javascript:submit()">
<option value="0">Selecione uma área</option>
<%
sql2="select area from grades_areacon_u where usuario='" & session("usuariomaster") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if areacon=rs2("area") then tempsel="selected" else tempsel=""
%>
	<option value="<%=rs2("area")%>" <%=tempsel%>><%=rs2("area")%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
<!-- <input type="submit" name="ver" value="Visualizar" class=button> -->
</form>
<%
coddoc=request("curso")
codmat=request("materia")
sql1="SELECT ap.area, m.materia " & _
"FROM grades_areacon_p2 ap INNER JOIN umaterias m ON ap.codmat=m.codmat " & _
"GROUP BY ap.area, m.materia " & _
"HAVING ap.area='" & areacon & "' ORDER BY m.materia;"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellspacing="2" cellpadding="1" style="border-collapse: collapse" width=500>
<%
do while not rs.eof
if codmat=rs("materia") then 
	imagem="minus.jpg" : status1=0
else
	imagem="plus.jpg"  : status1=1
end if
%>
<tr>
	<td width=15><a href="treed.asp?materia=<%if status1=1 then response.write rs("materia")%>">
	<img src="<%=imagem%>" border=0></a></td>
	<td colspan=3 class=campo><%=rs("materia")%></b></td>
</tr>
<%
if codmat=rs("materia") then
	sql2="SELECT ap.area, m.MATERIA, ap.chapa1, f.NOME, f.CODSITUACAO " & _
	"FROM UMATERIAS m INNER JOIN (grades_areacon_p2 ap INNER JOIN pfunc f ON ap.chapa1=f.CHAPA) ON m.CODMAT=ap.codmat " & _
	"WHERE m.MATERIA='" & codmat & "' " & _
	"GROUP BY ap.area, m.MATERIA, ap.chapa1, f.NOME, f.CODSITUACAO " & _
	"HAVING ap.area='" & areacon & "' "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
	if codmat=rs2("materia") then 
		imagem2="minus.jpg" : status2=0
	else 
		imagem2="plus.jpg"  : status2=1
	end if
	corletra="green"
	if rs2("codsituacao")="D" then corletra="red"
	if rs2("codsituacao")="A" then corletra="black"
	if rs2("codsituacao")="F" then corletra="black"
	if rs2("codsituacao")="Z" then corletra="black"
%>
<tr>
	<td width=15>&nbsp;</td>
	<td width=15>&nbsp;</td>
	<td width=15>
    <a class=r href="docente_ver.asp?chapa=<%=rs2("chapa1")%>" onclick="NewWindow(this.href,'CadastroProfessor','655','480','yes','center');return false" onfocus="this.blur()">	
	<img src="../images/bullet.gif" width="13" height="8" border="0" alt=""></a></td>
	<td width=100% colspan=1 class=campo>
    <a class=r href="docente_ver.asp?chapa=<%=rs2("chapa1")%>" onclick="NewWindow(this.href,'CadastroProfessor','655','480','yes','center');return false" onfocus="this.blur()">	
	<font color=<%=corletra%>><b><%=rs2("nome")%></a></td>
</tr>
<%
	rs2.movenext:loop
	rs2.close
end if
%>

<%
rs.movenext
loop
rs.close
%>
</table>

<%

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>