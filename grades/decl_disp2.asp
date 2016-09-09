<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
'if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
'if session("a80")="N" or session("a8")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Declaração de disponibilidade</title>
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
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B3")="" then
%>
<p class=titulo>Declaração de disponibilidade</p>
<form method="POST" action="decl_disp2.asp" name="form">
Professor:
	<select size="1" name="D1">
		<option value="Todos">Todos</option>
		<option value="Narciso">Todos-Campus Narciso</option>
		<option value="Yara">Todos-Campus V.Yara</option>
<%
sql1="select chapa, nome from dc_professor where codsituacao<>'D' order by nome "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
%>
<option value="<%=rs("chapa")%>"><%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
%>			
	</select>
<br><input type="submit" value="Clique para Visualizar" class="button" name="B3"></td></tr>
</form>

<%
end if 'request.form("B3")=""

if request.form("B3")<>"" then

temp=" and dc.chapa in (select chapa1 collate database_default from achapa) "
temp=" and dc.chapa>'00000' "

	chapa=request.form("d1")
	if chapa="Todos" then
		sql2=" " & temp
	elseif chapa="Narciso" then
		sql2=" and left(codsecao,2)='01' " & temp
	elseif chapa="Yara" then
		sql2=" and left(codsecao,2)='03' " & temp
	else
		sql2=" and chapa='" & chapa & "' "
	end if
	
	sql1="SELECT dc.CHAPA, dc.NOME, dc.DATAADMISSAO, dc.FUNCAO, dc.CODSECAO, s.descricao as SECAO, codsituacao " & _
"FROM dc_professor dc, corporerm.dbo.psecao s " & _
"WHERE dc.codsecao=s.codigo and dc.CODSITUACAO in ('A','F','Z','E','L') and codtipo='N' " & sql2 & " order by dc.codsecao, dc.nome  "

	'response.write sql1
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof
	ttaulas=0
	ttjornada=0
	ttsalario=0

'teste=1 imprime resumo
'teste=0 imprime cartas	
teste=0

if teste=0 then
%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<!--
<tr><td><img border="0" src="../images/logo_centro_universitario_unifieo_big.jpg" width=225></td><td width=35 rowspan=7></td></tr>
-->
<tr><td><p align="center">&nbsp;</td></tr>
<tr>
	<td>
	<br>Osasco, <%=day(now)%> de <%=monthname(month(now))%> de <%=year(now)%>
	<br>
    <br>
	<br>
	<br><b>A Fundação Instituto de Ensino para Osasco</b>
	<br>
    <br>
	<br>
	<br>Ref.: Declaração de disponibilidade de horários
	<br>
	<br>
	<br>
	<br>
    <p style="margin-bottom:0px;text-align:justify">
	Venho confirmar e declarar a minha disponibilidade de horários para o <b>primeiro semestre</b> letivo de 2014, conforme descrito abaixo:</p>
	
	<div align="center">
<%
sqlc="select d.chapa, f.nome, 'Dia'=case diasem when 2 then 'Segunda' when 3 then 'Terça' when 4 then 'Quarta' when 5 then 'Quinta' when 6 then 'Sexta' when 7 then 'Sábado' end " & _
",'07:30-08:20'= case m01 when 1 then 'X' else '' end ,'08:20-09:10'= case m02 when 1 then 'X' else '' end " & _
",'09:20-10:10'= case m03 when 1 then 'X' else '' end ,'10:10-11:00'= case m04 when 1 then 'X' else '' end " & _
",'11:10-12:00'= case m05 when 1 then 'X' else '' end ,'12:00-12:50'= case m06 when 1 then 'X' else '' end " & _
",'13:00-13:50'= case v01 when 1 then 'X' else '' end ,'13:50-14:40'= case v02 when 1 then 'X' else '' end " & _
",'14:50-15:40'= case v03 when 1 then 'X' else '' end ,'15:40-16:30'= case v04 when 1 then 'X' else '' end " & _
",'16:40-17:30'= case v05 when 1 then 'X' else '' end " & _
",'17:10-18:00'= case n01 when 1 then 'X' else '' end ,'18:00-18:50'= case n02 when 1 then 'X' else '' end " & _
",'19:00-19:50'= case n03 when 1 then 'X' else '' end ,'19:50-20:40'= case n04 when 1 then 'X' else '' end " & _
",'20:50-21:40'= case n05 when 1 then 'X' else '' end ,'21:40-22:30'= case n06 when 1 then 'X' else '' end " & _
"from grades_disp d inner join dc_professor f on f.chapa=d.chapa " & _
"where d.chapa='" & rs("chapa") & "' " 
rs2.Open sqlc, ,adOpenStatic, adLockReadOnly

'*************** inicio teste **********************
totaldisp=0
response.write "<table border='1' bordercolor='#000000' cellpadding='0' cellspacing='0' style='border-collapse:collapse' width='600'>"
response.write "<tr>"
for a= 2 to rs2.fields.count-1
	response.write "<td class=titulor>" & ucase(rs2.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs2.eof 
response.write "<tr>"
for a= 2 to rs2.fields.count-1
	if a>2 then alinhamento="center" else alinhamento="left"
	if rs2.fields(a)="X" then totaldisp=totaldisp+1
	response.write "<td class=""campor"" nowrap align='" & alinhamento & "'>" & rs2.fields(a) & "</td>"
next
response.write "</tr>"
rs2.movenext:loop
response.write "</table>"
'*************** fim teste **********************
rs2.close

sqlh="select chapa1, atual=[20121] from totalizador_2ch where chapa1='" & rs("chapa") & "' "
rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then atual=rs2("atual") else atual=0
rs2.close
%>
</div>
<%
if atual>20 then atual=20
if totaldisp<atual then 'situacao onde o professor disponibiliza menos horários do que a CH atual
%>
	<br>
    <p style="margin-bottom:0px;text-align:justify">
	Declaro ainda, estar ciente de que ao disponibilizar <%=totaldisp%> horários, menor do que minha jornada atual de <%=atual%> aulas semanais, estou abrindo 
	mão de <%=atual-totaldisp%> aulas semanais em razão de motivos:
	<br>(&nbsp;&nbsp;) pessoais
	<br>(&nbsp;&nbsp;) compromisso em outra instituição
	<br>(&nbsp;&nbsp;) _____________________________________
	</p>
<%
else
%>	
	<br>
    <p style="margin-bottom:0px;text-align:justify">
	Declaro ainda, que embora tenha disponibilizado um total de <%=totaldisp%> horários, estar ciente de que este número não é equivalente ao total de aulas 
	semanais para o próximo semestre, atualmente em <%if atual>20 then response.write "20" else response.write atual%> aulas semanais.
	</p>
<%
end if
%>
	
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr><td>Atenciosamente,    </td></tr>

<tr>
	<td>
	<p>&nbsp;
	<p>______________________________________________<br>
	<%=rs("nome")%> / <%=rs("chapa")%>
	<br>
	</td>
</tr>
<tr>
	<td class=campo>
	<br>
	<div align="center">
	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="580">
	<tr><td class="campor">
	Espaço reservado para anotações.
	<br><br><br><br><br><%=rs.absoluteposition%> - <%=rs("secao")%>

	</td></tr></table>
	</div

	</td>
</tr>
</table>
</div>
<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
end if 'teste 0
rs.movenext
loop

%>
</body>
</html>
<%


'*************** inicio teste **********************

if teste=1 then

response.write "<DIV style=""page-break-after:always""></DIV>"
response.write "<div align=""right"">"
rs.movefirst
response.write "<table border='1' bordercolor='#000000' cellpadding='0' cellspacing='0' style='border-collapse:collapse' width='650'>"
response.write "<tr>"
response.write "<td></td>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
response.write "<td>" & rs.absoluteposition & "</td>"
for a= 0 to rs.fields.count-1
	response.write "<td class=""campor"" nowrap>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext:loop
response.write "</table>"
response.write "</div>"

end if
'*************** fim teste **********************

rs.close


end if 'request.form("B3")<>""

set rs=nothing
set rs2=nothing

conexao.close
set conexao=nothing
%>