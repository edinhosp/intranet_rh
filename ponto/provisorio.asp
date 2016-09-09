<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")="N" or session("a48")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Provisórios</title>
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
dim conexao, conexao2, chapach, rs, rs2, tgl(4,6), tl(4), tg(6), descricao(4)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
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

dataatual=cdate(formatdatetime(now(),2))
datapesquisa=cdate(formatdatetime(request.form("dataquery"),2))
if request.form("dataquery")<>"" then
	datacampo=datapesquisa
	numero=2
else
	datacampo=dataatual
	numero=1
end if

sql1="select f.chapa, f.nome, u.codcracha, u.datainicio, u.datafim from corporerm.dbo.pfunc f, corporerm.dbo.ausoprov u where u.chapafunc=f.chapa " & _
"and getdate() between u.datainicio and u.datafim " & _
"order by u.codcracha, u.datainicio "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" name="form" action="provisorio.asp">
<p class=titulo>Controle de Crachás Provisórios

<table border=1>
<tr><td rowspan=4 height=100% valign="top">

<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=5>Crachás em poder de funcionários</td></tr>
<tr>
	<td class=titulo rowspan=2 align="center">Crachá</td>
	<td class=titulo rowspan=2 align="center">Chapa</td>
	<td class=titulo rowspan=2 align="center">Funcionário</td>
	<td class=titulo colspan=2 align="center">Período</td>
</tr>
<tr>
	<td class=titulo align="center">Inicio</td>
	<td class=titulo align="center">Final</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo align="center"><%=rs("codcracha")%>
	<a href="provisoriod_nova.asp?chapa=<%=rs("chapa")%>&provisorio=<%=rs("codcracha")%>" onclick="NewWindow(this.href,'provisorioD_nova','500','200','yes','center');return false" onfocus="this.blur()">
	<img src="../images/setanext1.gif" width="12" height="12" border="0" alt=""></a>
	</td>
	<td class=campo align="center"><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo align="center"><%=rs("datainicio")%></td>
	<td class=campo align="center"><%=rs("datafim")%></td>
</tr>
<%
rs.movenext
loop
else
end if
rs.close
%>
</table>

	</td>
	<td class=grupo align="center" height="15">Crachás Provisórios Entregues
	<a href="provisorio_nova.asp" onclick="NewWindow(this.href,'provisorioE_nova','500','200','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif"></a>
	</td>
</tr>
<tr>
	<td valign=top height="100%">
	<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td class=titulo>Provisório</td>
		<td class=titulo>Entregue a</td>
		<td class=titulo>Data</td>
		<td class=titulo>Hora</td>
		<td class=titulo>Por</td>
	</tr>
<%
sqlpe="select p.id_prov, p.operacao, p.provisorio, p.chapa, p.datae, p.horae, p.usuarioc, p.datac, f.nome " & _
"from provisorio p, corporerm.dbo.pfunc f where p.chapa=f.chapa collate database_default and operacao='E' order by p.provisorio "
rs.Open sqlpe, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
do while not rs.eof
hora=int(rs("horae")/60)
minuto=rs("horae")-(hora*60)
horae=numzero(hora,2) & ":" & numzero(minuto,2)
%>
	<tr>
		<td class=campo>
		<a href="provisorio_alteracao.asp?codigo=<%=rs("id_prov")%>" onclick="NewWindow(this.href,'ProvisorioE_alterar','500','200','yes','center');return false" onfocus="this.blur()">
		<%=rs("provisorio")%></a>
		</td>
		<td class=campo><%=rs("chapa")%>-<%=rs("nome")%></td>
		<td class=campo><%=rs("datae")%></td>
		<td class=campo><%=horae%></td>
		<td class=campo><%=rs("usuarioc")%></td>
	</tr>
<%
rs.movenext
loop
end if
rs.close
%>
	</table>	
	
	
	
	</td>
</tr>
<tr>
	<td class=grupo align="center" height="15">Crachás Provisórios Devolvidos
	<a href="provisoriod_nova.asp" onclick="NewWindow(this.href,'provisorioD_nova','500','200','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif"></a>
	</td>
</tr>
<tr>
	<td valign=top height="50%">
	<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td class=titulo>Provisório</td>
		<td class=titulo>Devolvido por</td>
		<td class=titulo>Data</td>
		<td class=titulo>Hora</td>
		<td class=titulo>Por</td>
	</tr>
<%
sqlpe="select p.id_prov, p.operacao, p.provisorio, p.chapa, p.datae, p.horae, p.usuarioc, p.datac, f.nome " & _
"from provisorio p, corporerm.dbo.pfunc f where p.chapa=f.chapa collate database_default and operacao='D' order by p.provisorio "
rs.Open sqlpe, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
do while not rs.eof
hora=int(rs("horae")/60)
minuto=rs("horae")-(hora*60)
horae=numzero(hora,2) & ":" & numzero(minuto,2)
%>
	<tr>
		<td class=campo>
		<a href="provisoriod_alteracao.asp?codigo=<%=rs("id_prov")%>" onclick="NewWindow(this.href,'ProvisorioD_alterar','500','200','yes','center');return false" onfocus="this.blur()">
		<%=rs("provisorio")%></a>
		</td>
		<td class=campo><%=rs("chapa")%>-<%=rs("nome")%></td>
		<td class=campo><%=rs("datae")%></td>
		<td class=campo><%=horae%></td>
		<td class=campo><%=rs("usuarioc")%></td>
	</tr>
<%
rs.movenext
loop
end if
rs.close
%>
	</table>	
	
	</td>
</tr>
</table>	

<br>

<%
sql2="select a.codcracha from corporerm.dbo.acracha a where a.situacao=1 and a.codcracha not in ( " & _
"select u.codcracha from corporerm.dbo.ausoprov u where getdate() between u.datainicio and u.datafim ) "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=2>Crachás Provisórios disponíveis</td></tr>
<tr><td class=titulo>Código do Crachá</td>
	<td class=titulo>Ultimo funcionário a usar</td>
</tr>
<%
rs.movefirst
do while not rs.eof
sql3="SELECT top 1 u.codcracha, u.datainicio, u.datafim, u.chapafunc fROM corporerm.dbo.AUSOPROV u WHERE u.codcracha='" & rs("codcracha") & "' " & _
" ORDER BY u.datafim DESC "
sql3="SELECT top 1 u.codcracha, u.datainicio, u.datafim, u.chapafunc, f.nome FROM corporerm.dbo.AUSOPROV u, corporerm.dbo.pfunc f WHERE u.codcracha='" & rs("codcracha") & "' " & _
"AND f.chapa=u.chapafunc ORDER BY u.datafim DESC "
rs2.Open sql3, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	uchapa=rs2("chapafunc"):unome=rs2("nome")
else
	uchapa="":unome=""
end if
rs2.close
%>
<tr>
	<td class=campo><%=rs("codcracha")%></td>
	<td class=campo><%=uchapa & " - " & unome%></td>
</tr>
<%
rs.movenext
loop
rs.close
%>
</table>



</form>
<%
	pagina=pagina+1
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
%>
<%

'rs.close
'set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>