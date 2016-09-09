<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a83")="N" or session("a83")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Saldo do Vale-Transporte</title>
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
	sql="SELECT vt_saldo.codigo, ptarifa.DESCRICAO, Sum([quantidade]*[fator]) AS quant, vt_saldo.tarifa, Sum([total]*[fator]) AS tot " & _
"FROM (vt_saldo INNER JOIN corporerm.dbo.ptarifa ptarifa ON vt_saldo.codigo = ptarifa.CODIGO collate database_default) INNER JOIN vt_saldo_tipo ON vt_saldo.id_tipo = vt_saldo_tipo.id_tipo " & _
"where deletada=0 and (getdate() between ptarifa.iniciovigencia and ptarifa.finalvigencia) " & _
"GROUP BY vt_saldo.codigo, ptarifa.DESCRICAO, vt_saldo.tarifa having Sum([total]*[fator])<>0"
'"where deletada=0 and (getdate() between ptarifa.iniciovigencia and ptarifa.finalvigencia) and vt_saldo.data='12/27/04' " & _
	rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="grupo" colspan="5">Controle de Saldo do Vale-Transporte - Saldo Atual</td>
</tr>
<tr>
	<td class="titulo">Código</td>
	<td class="titulo">Descrição</td>
	<td class="titulo">Quantidade</td>
	<td class="titulo">Tarifa</td>
	<td class="titulo">Total</td>
</tr>
<%
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class="campo"><%=rs("codigo")%></td>
	<td class="campo"><%=rs("descricao")%></td>
	<td class="campo" align="right"><%=formatnumber(rs("quant"),0)%>&nbsp;</td>
	<td class="campo" align="right"><%=formatnumber(rs("tarifa"),2)%>&nbsp;</td>
	<td class="campo" align="right"><%=formatnumber(rs("tot"),2)%>&nbsp;</td>
</tr>
<%
tgeral=tgeral+cdbl(rs("tot"))
rs.movenext
loop
%>
<tr>
	<td class="grupo" colspan="4">Saldo Total</td>
	<td class="titulo" align="right"><%=formatnumber(tgeral,2)%>&nbsp;</td>
</tr>
</table>
<% if session("grant_rh")="T" then %>
<a href="mov_nova.asp" onclick="NewWindow(this.href,'InclusaoVT','450','250','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo vt" WIDTH="16" HEIGHT="16">
<font size="1">inserir novo VT</font></a>
<% end if %>

<%
'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'rs.movefirst
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************
%>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>