<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a92")="N" or session("a92")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Grupo de Rotinas</title>
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

sqla="SELECT g.id_grupo, g.nome_grupo, Count(r.id_rotina) AS Quant " & _
"FROM rotinas_grupos g LEFT JOIN rotinas_0 r ON g.id_grupo=r.id_rotina " & _
"GROUP BY g.id_grupo, g.nome_grupo " & _
"ORDER BY g.nome_grupo "
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>Grupo de Rotinas

<table border="1" bordercolor="#000000" cellspacing="0" cellpadding="2" style="border-collapse: collapse" width="600">
  <tr>
    <td class=titulo align="center">Nome do Grupo</td>
    <td class=titulo align="center">Quant.<br>Rotinas</td>
    <td class=titulo align="center"><img border="0" src="../images/Magnify.gif"></td>
  </tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
%>
  <tr>
    <td class=campo><a href="#.asp?codigo=<%=rs("id_grupo")%>" class=r>
    <%=rs("nome_grupo")%></a></td>
    <td class=campo align="center"><%=rs("quant")%></td>
    <td class=campo align="center">&nbsp;
    <% if session("a92")="T" then %>
      <a href="grupo_alteracao.asp?codigo=<%=rs("id_grupo")%>" onclick="NewWindow(this.href,'AlteracaoGrupoRotina','520','100','no','center');return false" onfocus="this.blur()">
	  <img border="0" src="../images/folder95o.gif"></a>
	<% end if %>
    </td>
  </tr>
<%
rs.movenext
loop

else 'sem registros
%>
<tr><td colspan=3 class=grupo><b>Esta seleção não mostra nenhum registro.</b></td></tr>
<%
end if 'sem registros
%>
<tr><td colspan=3 class=titulo valign="center" align="right">
<% if session("a92")="T" then %>
<a href="grupo_nova.asp" onclick="NewWindow(this.href,'InclusaoGrupoRotina','520','100','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif">inserir novo grupo de rotinas</a>
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