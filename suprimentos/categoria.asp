<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a94")="N" or session("a94")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Grupo de Uniforme</title>
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
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

sqla="SELECT * from uniforme_categoria order by descricao "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>Categorias de Uniformes

<table border="1" cellspacing="0" cellpadding="2" style="border-collapse: collapse" width="600">
  <tr>
    <td class=titulo align="center">Categoria</td>
    <td class=titulo align="center"><img border="0" src="../images/Magnify.gif"></td>
  </tr>
<%
rs.movefirst
do while not rs.eof 
%>
  <tr>
    <td class=campo><a href="itens.asp?codigo=<%=rs("id_cat")%>" class=r>
    <%=rs("descricao")%></a></td>
    <td class=campo align="center">&nbsp;
    <% if session("a94")="T" then %>
      <a href="categ_alteracao.asp?codigo=<%=rs("id_cat")%>" onclick="NewWindow(this.href,'AlteracaoCategoria','520','150','no','center');return false" onfocus="this.blur()">
	  <img border="0" src="../images/folder95o.gif"></a>
	<% end if %>
    </td>
  </tr>
<%
rs.movenext
loop
%>
<tr><td colspan=5 class=titulo valign="center" align="right">
<% if session("a94")="T" then %>
<a href="categ_nova.asp" onclick="NewWindow(this.href,'InclusaoCategoria','520','150','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif">inserir nova categoria de uniforme</a>
<% end if %>
</td></tr>
</table>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>