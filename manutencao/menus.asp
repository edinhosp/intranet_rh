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
<title>Cadastro de Menus</title>
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
sql1="select m.idmenu, m.menu, m.sigla, m.descricao from intranet_menus m order by m.idmenu"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" name="form" action="menus.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Menus</p>

<table border="0" cellspacing="1" cellpadding="1" style="border-collapse: collapse" width="690px">
<tr>
	<td class="campo" align="left"></td>
	<td class="campo" align="right"></td>
</tr>
</table>

<table border="1" cellspacing="1" cellpadding="1" style="border-collapse: collapse" width="690px">
<tr>
	<td class="titulo">#</td>
	<td class="titulo">Nome do Menu</td>
	<td class="titulo">Sigla</td>
	<td class="titulo" width="17px" align="center">
		<a href="menus_nova.asp?idmenu=" onclick="NewWindow(this.href,'Menus_Nova','550','300','yes','center');return false" onfocus="this.blur()" Alt="Inclusão">
		<img src="../imagesr/page_new.gif" border="0" width="10" alt="Inclusão de Menu"></a>
	</td>
	<td class="campo" valign="top" width="60%"></td>
</tr>
<%
do while not rs.eof
valida=""
if trim(request("menu"))=trim(rs("idmenu")) then fundocel="fundo" else fundocel="campo"
%>
<tr>
	<td class="<%=fundocel%>" valign="top">
		<a class=r href="menus_alteracao.asp?idmenu=<%=rs("idmenu")%>" onclick="NewWindow(this.href,'Menus_Alterar','550','300','yes','center');return false" onfocus="this.blur()">
		<%=rs("idmenu")%></a>
	</td>
	<td class="<%=fundocel%>" valign="top"><%=rs("menu")%></td>
	<td class="<%=fundocel%>" valign="top"><%=rs("sigla")%></td>
	<td class="<%=fundocel%>" valign="top">
		<a class=r href="menus.asp?oper=submenu&menu=<%=rs("idmenu")%>">
	<img src="../imagesr/tables.gif" alt="A ser implementado"></a>
	</td>
	<td class="campo" valign="top">
<%
if request("oper")="submenu" and trim(request("menu"))=trim(rs("idmenu")) then
%>
	<table border="0" cellspacing="1" cellpadding="1" style="border-collapse: collapse" width="100%">
	<tr><td class="fundo" style="border-bottom:1px dotted black">#</td>
		<td class="fundo" style="border-bottom:1px dotted black">Descrição</td>
		<td class="fundo" style="border-bottom:1px dotted black">Link</td>
		<td class="fundo" style="border-bottom:1px dotted black">Sigla</td>
		<td class="fundo" style="border-bottom:1px dotted black" align="center">
			<a href="submenus_nova.asp?idmenu=<%=request("menu")%>&submenu=" onclick="NewWindow(this.href,'SubMenus_Nova','550','300','yes','center');return false" onfocus="this.blur()">
			<img src="../imagesr/page_new.gif" border="0" width="10" alt="Inclusão de Sub-Menu"></a>
		</td>
	</tr>
<%
	sql2="select s.idmenu, s.idsub, s.submenu, s.descricao, s.link, s.sigla, s.siglasub " & _
	"from intranet_submenus s where s.idmenu=" & request("menu") & " order by s.idmenu, s.idsub "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
%>
	<tr>
	<td class="campo" style="border-bottom:1px dotted black">
		<a class=r href="submenus_alteracao.asp?idmenu=<%=rs("idmenu")%>&submenu=<%=rs2("idsub")%>" onclick="NewWindow(this.href,'Menus_Alterar','550','300','yes','center');return false" onfocus="this.blur()">
		<%=rs2("idsub")%></a>
	</td>
	<td class="campo" style="border-bottom:1px dotted black"><%=rs2("descricao")%></td>
	<td class="campo" style="border-bottom:1px dotted black"><%=rs2("link")%></td>
	<td class="campo" style="border-bottom:1px dotted black"><%=rs2("siglasub")%></td>
	<td class="campo" style="border-bottom:1px dotted black">&nbsp;<td>
	</tr>
<%
	rs2.movenext
	loop
	rs2.close
%>
	</table>
<%
end if	
%>	
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