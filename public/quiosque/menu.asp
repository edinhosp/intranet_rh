<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<html>
<!-- Generated by AceHTML Freeware http://freeware.acehtml.com -->
<!-- Creation date: 30/10/03 -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Recursos Humanos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
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
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=no,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
</head>
<body style="margin-top: 0; margin-left: 0; margin-right:0;margin-bottom:0">
<!-- <img src="../images/quiosque.jpg" border="0" width="514" height="325"> -->
<%
session("quiosque")="02379"
if session("quiosque")="" then
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr><td class=campo valign="center" background="../images/q_top.jpg" width="502" height="26">&nbsp;<b>Usu�rio:</b> n�o logado</td></tr>
	<tr><td background="../images/q_corpow.jpg" width="502" height="248">
	Rotina de Login
	</td></tr>
	<tr><td class=campo valign="center" background="../images/q_barra.jpg" width="502" height="21">&nbsp;<%=now()%></td>
	</tr>
</table>
<%else%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr><td class=campo valign="center" background="../images/q_top.jpg" width="502" height="26">&nbsp;<b>Usu�rio:</b> <%=session("quiosque")%></td></tr>
	<tr><td background="../images/q_corpo.jpg" width="502" height="248">
<map name="menuq">
	<area shape="rect" coords="2,2,98,69" href="func_cadastro.asp" target="_blank" title="VisualizarDadosCadastrais">
	<area shape="rect" coords="110,1,202,69" href="teste1.asp">
	<area shape="rect" coords="222,0,313,67" href="teste2.asp">
</map>
<img src="../images/q_corpo.jpg" border="0" width="502" height="248" alt="" usemap="#menuq">
	</td></tr>
	<tr><td class=campo valign="center" background="../images/q_barra.jpg" width="502" height="21">&nbsp;<%=now()%></td>
	</tr>
</table>
<%end if%>
</body>
</html>