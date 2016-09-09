<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a55")="N" or session("a55")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Convênios com IES</title>
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

sqla="SELECT b.id_faculdade, b.Faculdade, b.Mantenedora, b.Endereco, b.CNPJ, b.Cidade, b.Email, b.Contato, b.telefone, Count(c.cursos) AS tcursos " & _
"FROM rhconveniobe b LEFT JOIN rhconveniobec c ON b.id_faculdade = c.id_faculdade " & _
"GROUP BY b.id_faculdade, b.Faculdade, b.Mantenedora, b.Endereco, b.CNPJ, b.Cidade, b.Email, b.Contato, b.telefone " & _
"ORDER BY b.Faculdade "
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>
Cadastro de Convênios com Instituições de Ensino
<table border="1" bordercolor="#CCCCCC" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulo align="center">Instituição</td>
	<td class=titulo align="center">Contato</td>
	<td class=titulo align="center">Telefone</td>
	<td class=titulo align="center">Email</td>
	<td class=titulo align="center">Cursos</td>
	<td class=titulo align="center"><img border="0" src="../images/Magnify.gif"></td>
</tr>
<%
rs.movefirst
do while not rs.eof 
%>
<tr>
	<td class=campo><%=rs("faculdade") %></td>
	<td class=campo><%=rs("contato") %></td>
	<td class=campo><%=rs("telefone") %></td>
	<td class=campo><%=rs("email") %></td>
	<td class=campo align="center">
	<a href="fac_cursos.asp?codigo=<%=rs("id_faculdade")%>">
	<font size=1><%=rs("tcursos")%></font></a>
	</td>
	<td class=campo align="center">
    <% if session("a55")="T" then %>
		<a href="fac_alteracao.asp?codigo=<%=rs("id_faculdade")%>" onclick="NewWindow(this.href,'Alteracao','520','300','no','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0" width="16" height="16" alt=""></a>
	<% end if %>
	</td>
</tr>
<%
rs.movenext
loop
%>
</table>
    <% if session("a55")="T" then %>
	<a href="fac_nova.asp" onclick="NewWindow(this.href,'Inclusao','520','300','no','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
	<font size="1">inserir nova instituição</font></a>
<%
end if
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>