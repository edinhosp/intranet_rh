<!-- #config timefmt="%m/%d/%y" -->
<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
'	Response.buffer=true
'	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a99")="N" or session("a99")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastro de Grupos</title>
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
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=yes';
win=window.open(mypage,myname,settings);}
// -->
</script>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
sql1="select g.idgrupo, g.descricao, g.sigla from intranet_grupos g order by g.idgrupo "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" name="form" action="grupos.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Grupos</p>

<table border="0" cellspacing="1" cellpadding="1" style="border-collapse: collapse" width="690px">
<tr>
	<td class="campo" align="left"></td>
	<td class="campo" align="right">
		<a href="grupos_nova.asp?codigo=" onclick="NewWindow(this.href,'Grupo_Nova','550','300','yes','center');return false" onfocus="this.blur()">
		<img src="../imagesr/page_new.gif" border="0" width="10" alt="Inclusão de Grupo"></a>
	</td>
</tr>
</table>

<table border="1" cellspacing="1" cellpadding="1" style="border-collapse: collapse" width="690px">
<tr>
	<td class="titulo">#</td>
	<td class="titulo">Nome do Grupo</td>
	<td class="titulo">Sigla</td>
	<td class="titulo"></td>
</tr>
<%
do while not rs.eof
valida=""
%>
<tr>
	<td class="campo">
		<a class=r href="grupos_alteracao.asp?codigo=<%=rs("idgrupo")%>" onclick="NewWindow(this.href,'Grupo_Alterar','550','300','yes','center');return false" onfocus="this.blur()">
		<%=rs("idgrupo")%></a>
	</td>
	<td class="campo"><%=rs("descricao")%></td>
	<td class="campo"><%=rs("sigla")%></td>
	<td class="campo"><img src="../imagesr/tables.gif" alt="A ser implementado">
	</td>
</tr>
<%
rs.movenext
loop
%>
</table>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>

<!-- -->
<!-- -->
</form>
</body>
</html>