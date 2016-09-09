<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Quiosque de Consulta - Recursos Humanos</title>
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
win=window.open(mypage,myname,settings);

}
// -->
</script>

</head>
<body background="../images/bgrh.jpg" bgproperties="fixed" scroll="no">
<script language="javascript" type="text/javascript">
NewWindow('menu.asp','MenuQuiosque','600','400','no','center');
</script>

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=100%>
<tr><td class=titulo>
<a href="menu.asp" onclick="NewWindow(this.href,'MenuQuiosque','600','400','no','center');return false;" onfocus="this.blur()" onmouseout="" onclick="javascript:this.close()">
<font size=6>Entrar</a>
</td></tr></table>

</body>
</html>
