<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")="N" or session("a72")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Espelho de marcações - Estagiários</title>
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
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1()		{form.chapa.value=form.nome.value;}
function chapa1()		{form.nome.value=form.chapa.value;}
function mand_ini1(muda) {
	temp=form.dtinigozo.value;
	inicio=new Date(temp.substr(6),temp.substr(3,2)-1,temp.substr(0,2));
	temp2=form.dtfimgozo.value;
	termino=new Date(temp2.substr(6),temp2.substr(3,2)-1,temp2.substr(0,2));
	dinicio=montharray[inicio.getMonth()]+" "+inicio.getDate()+", "+inicio.getFullYear()
	dfinal=montharray[termino.getMonth()]+" "+termino.getDate()+", "+termino.getFullYear()
	dias=(Math.round((Date.parse(dfinal)-Date.parse(dinicio))/(24*60*60*1000))*1)+1
	document.form.dias.value=dias
}
--></script>
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

sql="select top 1 * from est_parametro "
rs.Open sql, ,adOpenStatic, adLockReadOnly
ano=rs("ano")
mes=rs("mes")
descricao=rs("descricao")
inicio1=rs("inicio")
fim1=rs("fim")
limite=rs("limite")
rs.close

if request.form("inicio")="" then inicio=inicio1 else inicio=request.form("inicio")
if request.form("fim")="" then fim=fim1 else fim=request.form("fim")
if request.form("chapa")="" then chapa="" else chapa=request.form("chapa")
%>
<form method="POST" action="espelhoponto.asp" name="form">
<span style="font-size:11pt;font-weight:bold;">Espelho do Período
de <input type="text" size="8" name="inicio" value="<%=inicio%>" style="text-align:center" class="subli" onchange="form.submit();">
a <input type="text" size="8" name="fim" value="<%=fim%>" style="text-align:center" class="subli" onchange="form.submit();"></span> 
<table border="0" bordercolor=black cellpadding="2" cellspacing="1" style="border-collapse:collapse">
<tr>
	<td class=campo>
	<input type="text" value="<%=chapa%>" name="chapa" size="5" class="subli" onchange="chapa1();form.submit();" onfocus="javascript:window.status='Informe o chapa do funcionário'">
	<select size="1" name="nome" class="subli" onchange="nome1();form.submit();" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" >
<%
sql2="select chapa, nome from pfunc where codsituacao<>'D' and codtipo='T' order by nome "
sql2="select e.chapa, f.nome from corporerm.dbo.pfunc f, est_batfun e where e.chapa=f.chapa collate database_default and e.data between '" & dtaccess(inicio) & "' and '" & dtaccess(fim) & "' " & _
"group by e.chapa, f.nome "
'if session("dp_chapa")<>"" then sql2=sql2 & "and chapa='" & session("dp_chapa") & "'" else sql2=sql2 & "order by nome"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o funcionário....</option>"
rs.movefirst:do while not rs.eof
if chapa=rs("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rs("chapa")%>" <%=temp%>><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select>
	<a href="espelhomes.asp?chapa=<%=chapa%>" onclick="NewWindow(this.href,'Inclusao','650','450','yes','center');return false" onfocus="this.blur()">
	<img src="../images/espelho_imprimir.jpg" border="0" alt=""></a>

	</td>
</tr>
</table>
<table border="0" bordercolor=black cellpadding="0" cellspacing="0" style="border-collapse:collapse" width="100%">
<tr><td style="border-top:2 solid blue"></td></tr></table>
</form>
<%
if request.form<>"" then
	echapa=request.form("chapa")
	einicio=request.form("inicio")
	efim=request.form("fim")
%>
<table border="0" bordercolor=black cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="">
<tr>
	<td class=titulo colspan=2>&nbsp;Data</td>
	<td class=titulo colspan=2>&nbsp;Horário Cumprir</td>
	<td class=titulo colspan=3>&nbsp;Marcações</td>
	<td class=titulo colspan=2 align="right">H.Trab.</td>
	<td class=titulo colspan=1>&nbsp;Atraso</td>
	<td class=titulo colspan=1>&nbsp;Extra</td>
	<td class=titulo colspan=2>&nbsp;Ex.Aut.</td>
	<td class=titulo colspan=1>&nbsp;Falta</td>
	<td class=titulo colspan=1>&nbsp;</td>
</tr>
<%
	sql1="select * from est_batfun where chapa='" & echapa & "' and data between '" & dtaccess(einicio) & "' and '" & dtaccess(efim) & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	do while not rs.eof
	folga=0:texto=""
	if rs("feriado")>0 then folga=1:texto="<font color=red>&nbsp;FERIADO"
	if rs("descanso")>0 then folga=1:texto="<font color=red>&nbsp;DESCANSO"
%>
<tr>
	<td class=campo style="border-left: 1px solid;border-bottom: 1px solid #E3E3E3"><%=rs("data")%></td>
	<td class=campo style="border-right: 1px solid;border-bottom: 1px solid #E3E3E3"><%=weekdayname(weekday(rs("data")),2)%></td>
	<td class=campo style="background:transparent url('../images/1.gif') repeat fixed center;z-index:1;;border-bottom: 1px solid #E3E3E3" >
	<%if rs("feriado")>0 or rs("descanso")>0 then%>
		<%=texto%>
	<%else%>
	&nbsp;<%=horaload(rs("hor1"),2)%>
	&nbsp;<%=horaload(rs("hor2"),2)%>
	&nbsp;<%=horaload(rs("hor3"),2)%>
	&nbsp;<%=horaload(rs("hor4"),2)%>&nbsp;
	<%end if%>
	</td>
	<td class=campo style="border-left: 1px solid;border-right: 1px solid;border-bottom: 1px solid #E3E3E3">&nbsp;<%=horaload(rs("base"),2)%>&nbsp;</td>
	<td class=campo style="border-left: 1px solid;border-bottom: 1px solid #E3E3E3">
	<%if rs("ajust1")>0 and (rs("ajust1")-rs("marc1")<>0) then%>
		<span style="text-decoration:line-through;">&nbsp;<%=horaload(rs("marc1"),2)%>&nbsp;</span> 
	<%
		response.write "<font color=red>" & horaload(rs("ajust1"),2)
	else
		response.write "&nbsp;" & horaload(rs("marc1"),2)
	end if
	response.write "&nbsp;"
	%></td>
	<td class=campo style="border-left:1 dotted;border-bottom: 1px solid #E3E3E3">
	<%if rs("ajust2")>0 and (rs("ajust2")-rs("marc2")<>0) then%>
		<span style="text-decoration:line-through;">&nbsp;<%=horaload(rs("marc2"),2)%>&nbsp;</span> 
	<%
		response.write "<font color=red>" & horaload(rs("ajust2"),2)
	else
		response.write "&nbsp;" & horaload(rs("marc2"),2)
	end if
	response.write "&nbsp;"
	%></td>
	<td class=campo style="border-left:1 dotted;border-bottom: 1px solid #E3E3E3">
	<%if rs("ajust3")>0 and (rs("ajust3")-rs("marc3")<>0) then%>
		<span style="text-decoration:line-through;">&nbsp;<%=horaload(rs("marc3"),2)%>&nbsp;</span> 
	<%
		response.write "<font color=red>" & horaload(rs("ajust3"),2)
	else
		response.write "&nbsp;" & horaload(rs("marc3"),2)
	end if
	response.write "&nbsp;"
	%></td>
	<td class=campo style="border-left:1 dotted;border-bottom: 1px solid #E3E3E3">
	<%if rs("ajust4")>0 and (rs("ajust4")-rs("marc4")<>0) then%>
		<span style="text-decoration:line-through;">&nbsp;<%=horaload(rs("marc4"),2)%>&nbsp;</span> 
	<%
		response.write "<font color=red>" & horaload(rs("ajust4"),2)
	else
		response.write "&nbsp;" & horaload(rs("marc4"),2)
	end if
	response.write "&nbsp;"
	%></td>
	<td class=campo style="border-left: 1px solid;border-right: 1px solid;border-bottom: 1px solid #E3E3E3"><font color=blue>&nbsp;<%=horaload(rs("htrab"),2)%>&nbsp;</td>

	<td class=campo style="border-right: 1px solid;border-bottom: 1px solid #E3E3E3"><font color=red>&nbsp;<%=horaload(rs("atraso"),2)%>&nbsp;</td>
	<td class=campo style="border-right: 1px solid;border-bottom: 1px solid #E3E3E3"><font color=green>&nbsp;<%=horaload(rs("extra"),2)%>&nbsp;</td>
	<td class=campo style="border-right:0 solid;border-bottom: 1px solid #E3E3E3">&nbsp;<%=horaload(rs("extraaut"),2)%></td>
	<td class=campo style="border-right: 1px solid;border-bottom: 1px solid #E3E3E3" align="right">
<%if rs("extra")>0 then%>
		<a href="espelhoponto_extraaut.asp?chapa=<%=echapa%>&data=<%=rs("data")%>" onclick="NewWindow(this.href,'Alteracao','280','140','no','center');return false" onfocus="this.blur()">
		<img src="../images/espelho_extra.jpg" border="0" alt=""></a>
<%end if%>
	</td>
	<td class=campo style="border-right: 1px solid;border-bottom: 1px solid #E3E3E3">&nbsp;<%=horaload(rs("falta"),2)%>&nbsp;</td>
	<td class=campo style="border-right:1 dotted;border-bottom: 1px solid #E3E3E3" align="right">
		<a href="espelhoponto_alteracao.asp?chapa=<%=echapa%>" onclick="NewWindow(this.href,'Alteracao','510','350','yes','center');return false" onfocus="this.blur()">
		<img src="../images/espelho.gif" border="0" alt=""></a>
	</td>
</tr>
<%
	rs.movenext
	loop
%>
<tr><td colspan=14 style="border-top: 1px solid"></td></tr>
<tr>
	<td class=campo colspan=15 valign="top">
	<a href="espelhomes.asp?chapa=<%=echapa%>" onclick="NewWindow(this.href,'Inclusao','600','350','yes','center');return false" onfocus="this.blur()">
	<img src="../images/espelho_imprimir.jpg" border="0" alt=""><span style="color:black;font-weight:bold">Visualizar Espelho do Período</span>
	</a>
	</td>
</tr>
<tr><td colspan=15 style="border-top: 1px solid"></td></tr>

</table>
<%
end if 'request.form<>""

teste=0
	if teste=1 then
	'*************** inicio teste **********************
	if request.form<>"" then
	response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
	response.write "<tr>"
	for a=0 to rs.fields.count-1
		response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
	next
	response.write "</tr>"
	if rs.recordcount>0 then rs.movefirst
	do while not rs.eof 
	response.write "<tr>"
	for a= 0 to rs.fields.count-1
		response.write "<td class=""campor"" nowrap>" & rs.fields(a) & "</td>"
	next
	response.write "</tr>"
	rs.movenext
	loop
	response.write "</table>"
	response.write "<p>"
	end if
	'*************** fim teste **********************
end if 'teste=1
%>

</body>
</html>
<%

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>