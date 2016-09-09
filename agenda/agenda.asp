<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a96")="N" or session("a96")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Agenda de Compromissos e Lembretes</title>
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
<!-- -->
<table border="0" cellpadding="0" cellspacing="2" style="border-collapse: collapse">
<tr><td valign=top>
<!-- -->
<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open Application("consql")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsq=server.createobject ("ADODB.Recordset")
Set rsq.ActiveConnection = conexao2
usuario=session("usuariomaster")
if usuario="02379" or usuario="00892" or usuario="00259" or usuario="02159" or usuario="00595" or usuario="02463" or usuario="90261" or usuario="90344" then fotos=1 else fotos=0

dim emes(12),edia(31),fdia(40),edia2(31),edia3(31),fdia2(40),fdia3(40)
mesagora=month(now)
anoagora=year(now)
diaagora=day(now)

if request.form("visao")="1" then v1="checked":t="1" else v1=""
if request.form("visao")="2" then v2="checked":t="2" else v2=""
if request.form("visao")="3" then v3="checked":t="3" else v3=""
if v1="" and v2="" and v3="" then v1="checked":t="1"

if request("d")<>"" and request.form("diaagora")="" then diaagora=request("d") 
if request("m")<>"" and request.form("mesagora")="" then mesagora=request("m")
if request("a")<>"" and request.form("anoagora")="" then anoagora=request("a")
if request.form("diaagora")<>"" and request("d")="" then diaagora=request.form("diaform")
if request.form("mesagora")<>"" and request("m")="" then mesagora=request.form("mesform")
if request.form("anoagora")<>"" and request("a")="" then anoagora=request.form("anoform")
if request("t")<>"" then t=request("t")
if t="1" and v1="" then v1="checked"
if t="2" and v2="" then v2="checked"
if t="3" and v3="" then v3="checked"
		
if request.form<>"" then
	if request.form("B3")<>"" then
		finaliza=1
	else
		finaliza=0
		mesagora=request.form("mesform")
		diaagora=request.form("diaform")
		anoagora=request.form("anoform")
	end if
	if request.form("avanca")<>"" then
		mesagora=mesagora+1
		if mesagora>12 then
			mesagora=1
			anoagora=anoagora+1
		end if
	end if
	if request.form("volta")<>"" then
		mesagora=mesagora-1
		if mesagora<1 then
			mesagora=12
			anoagora=anoagora-1
		end if
	end if
	if request.form("avancay")<>"" then anoagora=anoagora+1
	if request.form("voltay")<>"" then anoagora=anoagora-1
end if

sqld="select day(diaferiado) as dia1 from gferiado " & _
"where month(diaferiado)=" & mesagora & " and year(diaferiado)=" & anoagora & " " & _
"group by day(diaferiado) "
rsq.Open sqld, ,adOpenStatic, adLockReadOnly
if rsq.recordcount>0 then
rsq.movefirst:do while not rsq.eof 
	edia(rsq("dia1"))=1
rsq.movenext:loop
end if
rsq.close

sqld="select day(a.data) as dia1 from agenda a " & _
"where month(a.data)=" & mesagora & " and year(a.data)=" & anoagora & " " & _
"group by day(a.data) "
sqld="select day(data) as dia1 from ( "
sqld=sqld & "SELECT a.data from agenda a where a.tipo=0 and a.usuarioc='" & session("usuariomaster") & "' and month(a.data)=" & mesagora & " and year(a.data)=" & anoagora & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a where a.tipo=2 and month(a.data)=" & mesagora & " and year(a.data)=" & anoagora & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a where a.tipo=3 and a.usuarioc='" & session("usuariomaster") & "' and month(a.data)=" & mesagora & " and year(a.data)=" & anoagora & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a, agenda_3 a3 where a3.id_agenda=a.id_agenda and a.tipo=3 and a3.codigo='" & session("usuariomaster") & "' and month(a.data)=" & mesagora & " and year(a.data)=" & anoagora & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a, usuarios u, agenda_1 a1 where a1.id_agenda=a.id_agenda and u.usuario=a.usuarioc and a.tipo=1 and a1.codigo=u.grupo and month(a.data)=" & mesagora & " and year(a.data)=" & anoagora & " "
sqld=sqld & ") as s group by day(data) "
rs.Open sqld, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof 
	fdia(rs("dia1"))=1
rs.movenext:loop
end if
rs.close

dianiv=diaagora
mesniv=mesagora
anoniv=anoagora
datasql=dateserial(anoagora,mesagora,diaagora)
if t="1" then incremento=0
if t="2" then incremento=7
if t="3" then incremento=30
datasql2=dateserial(anoagora,mesagora,diaagora+incremento)

sqla="select * from ("
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u where u.usuario=a.usuarioc and a.tipo=0 and a.usuarioc='" & session("usuariomaster") & "' and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u where u.usuario=a.usuarioc and a.tipo=2 and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u where u.usuario=a.usuarioc and a.tipo=3 and a.usuarioc='" & session("usuariomaster") & "' and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u, agenda_3 a3 where a3.id_agenda=a.id_agenda and u.usuario=a.usuarioc and a.tipo=3 and a3.codigo='" & session("usuariomaster") & "' and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u, agenda_1 a1 where a1.id_agenda=a.id_agenda and u.usuario=a.usuarioc and a.tipo=1 and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' and a1.codigo in (select grupo from usuarios where usuario='" & session ("usuariomaster") & "') "
sqla=sqla & ") as s order by data, hora "
'response.write sqla
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<!-- calencario -->
<%
diasemana=weekday(dateserial(anoagora,mesagora,1))
ultimodia=day(dateserial(anoagora,mesagora+1,1)-1)
ultimo=0
emes(1)="Janeiro":emes(2)="Fevereiro":emes(3)="Março":emes(4)="Abril":emes(5)="Maio":emes(6)="Junho"
emes(7)="Julho":emes(8)="Agosto":emes(9)="Setembro":emes(10)="Outubro":emes(11)="Novembro":emes(12)="Dezembro"
%>
<form method="POST" action="agenda.asp" name="form">

<input type="hidden" name="mesform" value="<%=mesagora%>">
<input type="hidden" name="diaform" value="<%=diaagora%>">
<input type="hidden" name="anoform" value="<%=anoagora%>">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="175">
<tr>
	<td class=campo><input type="submit" value="<<" name="voltay" class=button></td>
	<td class=campo><input type="submit" value="<" name="volta" class=button></td>
	<td class="campor" width="100%" align="center">
		<font color="#000080"><b><%=emes(mesagora)& "/" & anoagora%></font></td>
	<td class=campo><input type="submit" value=">" name="avanca" class=button></td>
	<td class=campo><input type="submit" value=">>" name="avancay" class=button></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="175">
<tr>
	<td class="campo" align="center">Dom</td>
	<td class="campo" align="center">Seg</td>
	<td class="campo" align="center">Ter</td>
	<td class="campo" align="center">Qua</td>
	<td class="campo" align="center">Qui</td>
	<td class="campo" align="center">Sex</td>
	<td class="campo" align="center">Sab</td>
</tr>
<tr>
<%
for linha=1 to 7
	'response.write "<td class=campo align='center'>"
	if linha=diasemana then
		ultimo=1:if fdia(ultimo)=1 then fundo="fundo" else fundo="campo"
		if edia(ultimo)=1 or linha=1 then 'é feriado
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora & "&t=" & t & "' class=r style='color:red'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		else
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora & "&t=" & t  & "' class=r>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if		
	elseif ultimo>=1 then
		ultimo=ultimo+1:if fdia(ultimo)=1 then fundo="fundo" else fundo="campo"
		if edia(ultimo)=1 then
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora & "&t=" & t  & "' class=r style='color:red'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		else
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora  & "&t=" & t   & "' class=r>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if
	else
		response.write "<td class=campo align='center'>"
	end if
	response.write "</td>"
next
response.write "</tr>"

vartemp1=ultimodia-ultimo
vartemp2=int(vartemp1/7)
if (vartemp1/7)-vartemp2>0 then vartemp2=vartemp2+1
for sem=1 to vartemp2
	response.write "<tr>"
	for l2=1 to 7
		ultimo=ultimo+1:if fdia(ultimo)=1 then fundo="fundo" else fundo="campo"
		response.write "<td class=" & fundo & " align='center'>"
		if ultimo<=ultimodia then 
			if edia(ultimo)=1 or l2=1 then
				response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora  & "&t=" & t   & "' class=r style='color:red'>"
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</a>"
			else
				response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora  & "&t=" & t   & "' class=r>"
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</a>"
			end if
		end if
		response.write "</td>"
	next
	response.write "</tr>"
next
%>
</table>
<!-- fim calencario -->

<!-- segundo mês - INICIO INICIO INICIO -->
<%
anoagora2=anoagora
diasemana=weekday(dateserial(anoagora,mesagora+1,1))
ultimodia=day(dateserial(anoagora,mesagora+2,1)-1)
ultimo=0
mesagora2=mesagora+1
if mesagora2>12 then mesagora2=1:anoagora2=anoagora+1
sqld="select day(diaferiado) as dia1 from gferiado " & _
"where month(diaferiado)=" & mesagora2 & " and year(diaferiado)=" & anoagora2 & " " & _
"group by day(diaferiado) "
rsq.Open sqld, ,adOpenStatic, adLockReadOnly
if rsq.recordcount>0 then
rsq.movefirst:do while not rsq.eof 
	edia2(rsq("dia1"))=1
rsq.movenext:loop
end if
rsq.close

sqld="select day(data) as dia1 from ( "
sqld=sqld & "SELECT a.data from agenda a where a.tipo=0 and a.usuarioc='" & session("usuariomaster") & "' and month(a.data)=" & mesagora2 & " and year(a.data)=" & anoagora2 & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a where a.tipo=2 and month(a.data)=" & mesagora2 & " and year(a.data)=" & anoagora2 & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a where a.tipo=3 and a.usuarioc='" & session("usuariomaster") & "' and month(a.data)=" & mesagora2 & " and year(a.data)=" & anoagora2 & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a, agenda_3 a3 where a3.id_agenda=a.id_agenda and a.tipo=3 and a3.codigo='" & session("usuariomaster") & "' and month(a.data)=" & mesagora2 & " and year(a.data)=" & anoagora2 & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a, usuarios u, agenda_1 a1 where a1.id_agenda=a.id_agenda and u.usuario=a.usuarioc and a.tipo=1 and a1.codigo=u.grupo and month(a.data)=" & mesagora2 & " and year(a.data)=" & anoagora2 & " "
sqld=sqld & ") as s group by day(data) "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof 
	fdia2(rs2("dia1"))=1
rs2.movenext:loop
end if
rs2.close

%>
<hr>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="175">
<tr>
	<td class="campor" width="100%" align="center">
		<font color="#000080"><b><%=emes(mesagora2)& "/" & anoagora2%></font></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="175">
<tr>
	<td class="campo" align="center">Dom</td>
	<td class="campo" align="center">Seg</td>
	<td class="campo" align="center">Ter</td>
	<td class="campo" align="center">Qua</td>
	<td class="campo" align="center">Qui</td>
	<td class="campo" align="center">Sex</td>
	<td class="campo" align="center">Sab</td>
</tr>
<tr>
<%
for linha=1 to 7
	'response.write "<td class=campo align='center'>"
	if linha=diasemana then
		ultimo=1:if fdia2(ultimo)=1 then fundo="fundo" else fundo="campo"
		if edia2(ultimo)=1 or linha=1 then 'é feriado
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora2 & "&a=" & anoagora2  & "&t=" & t   & "' class=r style='color:red'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		else
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora2 & "&a=" & anoagora2  & "&t=" & t   & "' class=r>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if		
	elseif ultimo>=1 then
		ultimo=ultimo+1:if fdia2(ultimo)=1 then fundo="fundo" else fundo="campo"
		if edia2(ultimo)=1 then
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora2 & "&a=" & anoagora2  & "&t=" & t   & "' class=r style='color:red'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		else
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora2 & "&a=" & anoagora2  & "&t=" & t   & "' class=r>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if
	else
		response.write "<td class=campo align='center'>"
	end if
	response.write "</td>"
next
response.write "</tr>"

vartemp1=ultimodia-ultimo
vartemp2=int(vartemp1/7)
if (vartemp1/7)-vartemp2>0 then vartemp2=vartemp2+1
for sem=1 to vartemp2
	response.write "<tr>"
	for l2=1 to 7
		ultimo=ultimo+1:if fdia2(ultimo)=1 then fundo="fundo" else fundo="campo"
		response.write "<td class=" & fundo & " align='center'>"
		if ultimo<=ultimodia then 
			if edia2(ultimo)=1 or l2=1 then
				response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora2 & "&a=" & anoagora2  & "&t=" & t   & "' class=r style='color:red'>"
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</a>"
			else
				response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora2 & "&a=" & anoagora2  & "&t=" & t   & "' class=r>"
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</a>"
			end if
		end if
		response.write "</td>"
	next
	response.write "</tr>"
next
%>
</table>
<!-- segundo mês - FIM FIM FIM -->

<!-- terceiro mês - INICIO INICIO INICIO -->
<%
anoagora3=anoagora2
diasemana=weekday(dateserial(anoagora,mesagora+2,1))
ultimodia=day(dateserial(anoagora,mesagora+3,1)-1)
ultimo=0
mesagora3=mesagora2+1
if mesagora3>12 then mesagora3=1:anoagora3=anoagora2+1
sqld="select day(diaferiado) as dia1 from gferiado " & _
"where month(diaferiado)=" & mesagora3 & " and year(diaferiado)=" & anoagora3 & " " & _
"group by day(diaferiado) "
rsq.Open sqld, ,adOpenStatic, adLockReadOnly
if rsq.recordcount>0 then
rsq.movefirst:do while not rsq.eof 
	edia3(rsq("dia1"))=1
rsq.movenext:loop
end if
rsq.close

sqld="select day(data) as dia1 from ( "
sqld=sqld & "SELECT a.data from agenda a where a.tipo=0 and a.usuarioc='" & session("usuariomaster") & "' and month(a.data)=" & mesagora3 & " and year(a.data)=" & anoagora3 & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a where a.tipo=2 and month(a.data)=" & mesagora3 & " and year(a.data)=" & anoagora3 & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a where a.tipo=3 and a.usuarioc='" & session("usuariomaster") & "' and month(a.data)=" & mesagora3 & " and year(a.data)=" & anoagora3 & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a, agenda_3 a3 where a3.id_agenda=a.id_agenda and a.tipo=3 and a3.codigo='" & session("usuariomaster") & "' and month(a.data)=" & mesagora3 & " and year(a.data)=" & anoagora3 & " "
sqld=sqld & "union all "
sqld=sqld & "SELECT a.data from agenda a, usuarios u, agenda_1 a1 where a1.id_agenda=a.id_agenda and u.usuario=a.usuarioc and a.tipo=1 and a1.codigo=u.grupo and month(a.data)=" & mesagora3 & " and year(a.data)=" & anoagora3 & " "
sqld=sqld & ") as s group by day(data) "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof 
	fdia3(rs2("dia1"))=1
rs2.movenext:loop
end if
rs2.close
%>
<hr>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="175">
<tr>
	<td class="campor" width="100%" align="center">
		<font color="#000080"><b><%=emes(mesagora3)& "/" & anoagora3%></font></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="175">
<tr>
	<td class="campo" align="center">Dom</td>
	<td class="campo" align="center">Seg</td>
	<td class="campo" align="center">Ter</td>
	<td class="campo" align="center">Qua</td>
	<td class="campo" align="center">Qui</td>
	<td class="campo" align="center">Sex</td>
	<td class="campo" align="center">Sab</td>
</tr>
<tr>
<%
for linha=1 to 7
	'response.write "<td class=campo align='center'>"
	if linha=diasemana then
		ultimo=1:if fdia3(ultimo)=1 then fundo="fundo" else fundo="campo"
		if edia3(ultimo)=1 or linha=1 then 'é feriado
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora3 & "&a=" & anoagora3  & "&t=" & t   & "' class=r style='color:red'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		else
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora3 & "&a=" & anoagora3  & "&t=" & t   & "' class=r>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if		
	elseif ultimo>=1 then
		ultimo=ultimo+1:if fdia3(ultimo)=1 then fundo="fundo" else fundo="campo"
		if edia3(ultimo)=1 then
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora3 & "&a=" & anoagora3  & "&t=" & t   & "' class=r style='color:red'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		else
			response.write "<td class=" & fundo & " align='center'>"
			response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora3 & "&a=" & anoagora3  & "&t=" & t   & "' class=r>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if
	else
		response.write "<td class=campo align='center'>"
	end if
	response.write "</td>"
next
response.write "</tr>"

vartemp1=ultimodia-ultimo
vartemp2=int(vartemp1/7)
if (vartemp1/7)-vartemp2>0 then vartemp2=vartemp2+1
for sem=1 to vartemp2
	response.write "<tr>"
	for l2=1 to 7
		ultimo=ultimo+1:if fdia3(ultimo)=1 then fundo="fundo" else fundo="campo"
		response.write "<td class=" & fundo & " align='center'>"
		if ultimo<=ultimodia then 
			if edia3(ultimo)=1 or l2=1 then
				response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora3 & "&a=" & anoagora3  & "&t=" & t   & "' class=r style='color:red'>"
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</a>"
			else
				response.write "<a href='agenda.asp?d=" & ultimo & "&m=" & mesagora3 & "&a=" & anoagora3  & "&t=" & t   & "' class=r>"
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</a>"
			end if
		end if
		response.write "</td>"
	next
	response.write "</tr>"
next
%>
</table>
<!-- terceiro mês - FIM FIM FIM -->
<hr>
<input type="radio" name="visao" value="1" <%=v1%> onClick="javascrip:submit()"> dia
<input type="radio" name="visao" value="2" <%=v2%> onClick="javascrip:submit()"> semana
<input type="radio" name="visao" value="3" <%=v3%> onClick="javascrip:submit()"> mês

</form>

<!-- -->
</td>
<td valign=top>
<!-- -->
<% 
%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="400">
<%if t="1" then%>
<tr><td colspan=4 class="campol" align="center"><font size="2">&nbsp;Lembretes/Compromissos do dia <b><%=datasql%>&nbsp;</font></td></tr>
<tr><td class=fundo align="center"></td>
	<td class=fundo align="center" valign=middle width=35>Hora</td>
	<td class=fundo align="center" valign=middle width=295>Compromisso</td>
	<td class=fundo align="center" valign=middle width=70>Anotação<br>criada por:</td>
</tr>
<%end if%>
<%if t="2" or t="3" then%>
<tr><td colspan=5 class="campol" align="center"><font size="2">&nbsp;Lembretes/Compromissos entre o dia <b><%=datasql%> e <%=datasql2%>&nbsp;</font></td></tr>
<tr><td class=fundo align="center"></td>
	<td class=fundo align="center" valign=middle width=50>Data</td>
	<td class=fundo align="center" valign=middle width=35>Hora</td>
	<td class=fundo align="center" valign=middle width=245>Compromisso</td>
	<td class=fundo align="center" valign=middle width=70>Anotação<br>criada por:</td>
</tr>
<%end if%>

<%
if rs.recordcount>0 then
	rs.movefirst:do while not rs.eof 
%>
<%if t="1" then%>
<tr><td class=campo align="center" height=16>
	<%if rs("usuarioc")=session("usuariomaster") then%>
	<a href="agenda_alteracao.asp?codigo=<%=rs("id_agenda")%>" onclick="NewWindow(this.href,'Alteracao','490','300','no','center');return false" onfocus="this.blur()">
	<img src="../images/LeafSearch.gif" width="16" height="16" border="0" alt=""></a>
	<%end if%>
	</td>
	<td class=campo align="center"><%if rs("hora")<>"" then response.write formatdatetime(rs("hora"),4) else response.write "-"%></td>
	<td class=campo>&nbsp;<%=rs("compromisso")%></td>
	<td class=campo>&nbsp;<%=rs("nome")%></td></tr>
<%end if%>
<%if t="2" or t="3" then
data1=numzero(day(rs("data")),2) & "/" & numzero(month(rs("data")),2) & " "
%>
<tr><td class=campo align="center" height=16>
	<%if rs("usuarioc")=session("usuariomaster") then%>
	<a href="agenda_alteracao.asp?codigo=<%=rs("id_agenda")%>" onclick="NewWindow(this.href,'Alteracao','490','300','no','center');return false" onfocus="this.blur()">
	<img src="../images/LeafSearch.gif" width="16" height="16" border="0" alt=""></a>
	<%end if%>
	</td>
	<td class=campo align="center">&nbsp;<%=data1%>&nbsp;</td>
	<td class=campo align="center"><%if rs("hora")<>"" then response.write formatdatetime(rs("hora"),4) else response.write "-"%></td>
	<td class=campo>&nbsp;<%=rs("compromisso")%></td>
	<td class=campo>&nbsp;<%=rs("nome")%></td></tr>
<%end if%>
<%
rs.movenext:loop
else
	if t="1" then response.write "<tr><td colspan='4'><font size='2'>&nbsp;Não há compromissos no dia de hoje.&nbsp;&nbsp;</font></td></tr>"
	if t="2" or t="3" then response.write "<tr><td colspan='5'><font size='2'>&nbsp;Não há compromissos neste período.&nbsp;&nbsp;</font></td></tr>"
end if 'if recordcount
%>
</table>
<a href="agenda_nova.asp?data=<%=datasql%>" onclick="NewWindow(this.href,'Inclusao','490','300','no','center');return false" onfocus="this.blur()">
<img src="../images/Appointment.gif" width="16" height="16" border="0" alt=""><font size="1">incluir nova anotação</font></a>
<%
rs.close

set rs=nothing
conexao.close
set conexao=nothing
%>

<!-- -->
</td></tr></table>
<!-- -->
</body>
</html>