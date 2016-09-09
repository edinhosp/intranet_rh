<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Cadastro de Grade Horária</title>
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
<style type="text/css">
#dhtmltooltip{
position: absolute;
left: -300px;
width: 150px;
border: 1px solid black;
padding: 2px;
background-color: lightyellow;
visibility: hidden;
z-index: 100;
/*Remove below line to remove shadow. Below line should always appear last within this CSS*/
filter: progid:DXImageTransform.Microsoft.Shadow(color=gray,direction=135);
}
#dhtmlpointer{
position:absolute;
left: -300px;
z-index: 101;
visibility: hidden;
}
</style>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function periodo1() {	form.perlanc.value="Todos";	}
function perlanc1() {	form.periodo.value="Todos";	}
--></script>
</head>
<body>
<script type="text/javascript">
/***********************************************
* Cool DHTML tooltip script II- © Dynamic Drive DHTML code library (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
***********************************************/
var offsetfromcursorX=12 //Customize x offset of tooltip
var offsetfromcursorY=10 //Customize y offset of tooltip

var offsetdivfrompointerX=10 //Customize x offset of tooltip DIV relative to pointer image
var offsetdivfrompointerY=14 //Customize y offset of tooltip DIV relative to pointer image. Tip: Set it to (height_of_pointer_image-1).

document.write('<div id="dhtmltooltip"></div>') //write out tooltip DIV
document.write('<img id="dhtmlpointer" src="../images/arrow2.gif">') //write out pointer image

var ie=document.all
var ns6=document.getElementById && !document.all
var enabletip=false
if (ie||ns6)
var tipobj=document.all? document.all["dhtmltooltip"] : document.getElementById? document.getElementById("dhtmltooltip") : ""

var pointerobj=document.all? document.all["dhtmlpointer"] : document.getElementById? document.getElementById("dhtmlpointer") : ""

function ietruebody(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

function ddrivetip(thetext, thewidth, thecolor){
if (ns6||ie){
if (typeof thewidth!="undefined") tipobj.style.width=thewidth+"px"
if (typeof thecolor!="undefined" && thecolor!="") tipobj.style.backgroundColor=thecolor
tipobj.innerHTML=thetext
enabletip=true
return false
}
}

function positiontip(e){
if (enabletip){
var nondefaultpos=false
var curX=(ns6)?e.pageX : event.clientX+ietruebody().scrollLeft;
var curY=(ns6)?e.pageY : event.clientY+ietruebody().scrollTop;
//Find out how close the mouse is to the corner of the window
var winwidth=ie&&!window.opera? ietruebody().clientWidth : window.innerWidth-20
var winheight=ie&&!window.opera? ietruebody().clientHeight : window.innerHeight-20

var rightedge=ie&&!window.opera? winwidth-event.clientX-offsetfromcursorX : winwidth-e.clientX-offsetfromcursorX
var bottomedge=ie&&!window.opera? winheight-event.clientY-offsetfromcursorY : winheight-e.clientY-offsetfromcursorY

var leftedge=(offsetfromcursorX<0)? offsetfromcursorX*(-1) : -1000

//if the horizontal distance isn't enough to accomodate the width of the context menu
if (rightedge<tipobj.offsetWidth){
//move the horizontal position of the menu to the left by it's width
tipobj.style.left=curX-tipobj.offsetWidth+"px"
nondefaultpos=true
}
else if (curX<leftedge)
tipobj.style.left="5px"
else{
//position the horizontal position of the menu where the mouse is positioned
tipobj.style.left=curX+offsetfromcursorX-offsetdivfrompointerX+"px"
pointerobj.style.left=curX+offsetfromcursorX+"px"
}

//same concept with the vertical position
if (bottomedge<tipobj.offsetHeight){
tipobj.style.top=curY-tipobj.offsetHeight-offsetfromcursorY+"px"
nondefaultpos=true
}
else{
tipobj.style.top=curY+offsetfromcursorY+offsetdivfrompointerY+"px"
pointerobj.style.top=curY+offsetfromcursorY+"px"
}
tipobj.style.visibility="visible"
if (!nondefaultpos)
pointerobj.style.visibility="visible"
else
pointerobj.style.visibility="hidden"
}
}

function hideddrivetip(){
if (ns6||ie){
enabletip=false
tipobj.style.visibility="hidden"
pointerobj.style.visibility="hidden"
tipobj.style.left="-1000px"
tipobj.style.backgroundColor=''
tipobj.style.width=''
}
}

document.onmousemove=positiontip
</script>
<%
dim conexao, rs, rs2, rs3
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
'response.write "<br> Request: " & request.form
'response.write "<br>Antes: " & "per: " & session("per101")& " cur: "  & session("cur101") & " tur: "  & session("tur101")

if request.form("periodo")="" then session("per101")="Todos" else session("per101")=request.form("periodo")
	if session("per101")<>"Todos" then
		session("sqlp101")=" AND perlet='" & session("per101") & "' "
		sqlp101=" AND perlet='" & session("per101") & "' "
	else
		session("sqlp101")=""
		sqlp101=""
	end if
if request.form("curso")="" then session("cur101")="Todos" else session("cur101")=request.form("curso")
	if session("cur101")<>"Todos" then
		session("sqlp101")=" AND (coddoc='" & session("cur101") & "') "
		sqlc101="AND (coddoc='" & session("cur101") & "') "
	else
		session("sqlc101")=""
		sqlc101=""
	end if
if request.form("turma")="" then session("tur101")="Todas" else session("tur101")=request.form("turma")
	if session("tur101")<>"Todas" then
		session("sqlt101")=" AND (turma='" & session("tur101") & "') "
		sqlt101=" AND (turma='" & session("tur101") & "') "
	else
		session("sqlt101")=""
		sqlt101=""
	end if
'response.write "<br>Depois: " & "per: " & session("per101")& " cur: "  & session("cur101") & " tur: "  & session("tur101")

sessao=session("usuariomaster")
terminoperiodo=now()
inicio=now()
%>
<form method="POST" action="gradesv3.asp" name="form">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Grade Horária</p>
<table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" width="690">
<tr><td class=campo></td></tr></table>

<!-- Periodo letivo -->
&nbsp;&nbsp;Período Letivo: <select size="1" name="periodo" onchange="javascript:submit()">
<option value="Todos" <%if session("per101")="Todos" then response.write "selected"%>>Todos</option>
<%
sql2="select perlet from g2turmas where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') group by perlet order by perlet desc"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("perlet")%>" <%if session("per101")=rs2("perlet") then response.write "selected"%>><%=rs2("perlet")%></option>
<%
rs2.movenext:loop
end if
rs2.close
%>
</select>

&nbsp;&nbsp;Filtrar Curso: <select size="1" name="curso" onchange="javascript:submit()">
<option value="Todos" <%if session("cur101")="Todos" then response.write "selected"%>>Todos cursos</option>
<%
sql2="select t.coddoc, c.curso from g2turmas t, g2cursos c " & _
"where c.coddoc=t.coddoc and c.codcur=t.codcur and c.codper=t.codper and t.coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') " & _
"" & sqlp101 & _
"group by t.coddoc, c.curso order by c.curso "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("coddoc")%>" <%if session("cur101")=rs2("coddoc") then response.write "selected"%>><%=rs2("curso")%></option>
<%
rs2.movenext:loop
end if
rs2.close
%>
</select>

<!-- turmas -->
&nbsp;&nbsp;Turma: <select size="1" name="turma" onchange="javascript:submit()">
<option value="Todas" <%if session("tur101")="Todas" then response.write "selected"%>>Todas</option>
<%
sql2="select codtur, serie, turma from g2turmas where coddoc='" & request.form("curso") & "' and perlet='" & request.form("periodo") & "' order by serie, turma, codtur"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("codtur")%>" <%if session("tur101")=rs2("codtur") then response.write "selected"%>><%=rs2("codtur")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>
<br>

<%
perlet=request.form("periodo")
coddoc=request.form("curso")
codtur=request.form("turma")
sql2="delete from g2temp where sessao='" & sessao & "'":conexao.execute sql2
codcur=0:codper=0:grade=0:serie=0:codhor=0:id_grdturma=0
sql1="select * from g2turmas where perlet='" & perlet & "' and coddoc='" & coddoc & "' and codtur='" & codtur & "' "
sql1="select t.*, c.tipocurso, inicio, termino, lancamento from g2turmas t, g2cursos c, g2periodoaula p where p.perlet=t.perlet and t.codcur=c.codcur and t.codper=c.codper and " & _
"t.perlet='" & perlet & "' and t.coddoc='" & coddoc & "' and t.codtur='" & codtur & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
'*************** inicio teste **********************
if session("usuariomaster")="02379" then
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext:loop
response.write "</table>"
response.write "# " & rs.recordcount & "<br>"
end if
'*************** fim teste **********************
if rs.recordcount>0 then 
	rs.movefirst
	codcur=rs("codcur"):codper=rs("codper"):grade=rs("grade"):serie=rs("serie"):tipocurso=rs("tipocurso"):turno=rs("turno")
	aberto=rs("aberto"):turmaid=rs("turmaid")
	id_grdturma=rs("id_grdturma")
	if tipocurso="2" then
	sql3="insert into g2temp (sessao, descricao, horini, horfim) select '" & sessao & "', descricao, horini, horfim from g2defhor " & _
	"where codtn=" & turno & " and tipocurso=" & tipocurso & " group by descricao, horini, horfim"
	conexao.execute sql3
	for a=2 to 7
		sql4="update g2temp set [" & a & "]=d.codhor from g2temp t, g2defhor d " & _
		"where d.codtn=" & turno & " and d.tipocurso=" & tipocurso & " and codds=" & a & " and d.horini=t.horini and d.horfim=t.horfim "
		conexao.execute sql4
		sql5="update g2temp set [" & a & "d]=1 from g2temp t, g2disp d " & _
		"where d.turno=" & turno & " and d.diasem=" & a & " and d.horini=t.horini and d.horfim=t.horfim and d.chapa='" & right(request.form("seldisc"),5) & "'"
		if right(request.form("seldisc"),5)<>"" then conexao.execute sql5
	next
	end if 'tipocurso=2
	inicioperiodo=rs("inicio"):iniciocalendar=dateserial(year(inicioperiodo),month(inicioperiodo),1)
	terminoperiodo=rs("termino")
	periodolanc=rs("lancamento")
	terminoanterior=inicioperiodo-1
	terminoanterior=year(terminoanterior)&numzero(month(terminoanterior),2)&numzero(day(terminoanterior),2)
	if aberto<>0 then msgaberto="<font color=blue>Alterações liberadas até " & periodolanc-1 & ".</font>" else msgaberto="<font color=red>Alterações Encerradas.</font>"
	semturma=0
else
	semturma=1
end if
if session("usuariomaster")="02379" then response.write request.form
rs.close


%>
<!-- -->
<table border="0" bordercolor="#000000" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="">
<tr><td valign=top style="border-right:2px solid black;border-top:2px solid black" colspan=2>
<!-- -->
<%
response.write msgaberto
if tipocurso="2" or tipocurso="12" then
%>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="">
<tr>
	<td class=titulo>Horários</td>
	<td class=título>Segunda-feira</td>
	<td class=título>Terça-feira</td>
	<td class=título>Quarta-feira</td>
	<td class=título>Quinta-feira</td>
	<td class=título>Sexta-feira</td>
	<td class=título>Sábado</td>
</tr>
<%
sql6="select * from g2temp where sessao='" & sessao & "' order by descricao "
rs.Open sql6, ,adOpenStatic, adLockReadOnly
do while not rs.eof
if cor="blue" then cor="red" else cor="blue"
%>
<tr>
	<td class=campo><font color=<%=cor%>><%=rs("horini")%></font>-<%=rs("horfim")%></td>
<%
for b=2 to 7
	texto="":nome1="":nome2=""
	codhor=rs("" & b & "")
	'response.write codhor
	if codhor>0 then
		sql72="select id_grdaula, g.codmat, m.materia, g.chapa1, chapa2, inicio, termino, juntar, jturma, dividir, dturma " & _
		"from g2aulas g, corporerm.dbo.umaterias m where g.codmat=m.codmat collate database_default " & _ 
		"and g.id_grdturma=" & id_grdturma & " and g.codhor=" & codhor & " and g.deletada=0 " & _
		"order by termino, inicio"
		sql7=sql72
		'response.write sql7
		rs2.Open sql7, ,adOpenStatic, adLockReadOnly
		registros=rs2.recordcount

		'*********** checar se o horario está disponivel
		if rs("" & b & "d")="1" then
			disponivel=1
			borda=" style='border:2px solid blue;background:#ccffcc'"
		else 
			disponivel=0 
			borda=" style='border:1px solid red;'"
		end if
		
		'response.write "<br>" & codhor & "-" & registros
		risc1="<font style='text-decoration:line-through'>"
		risc2="<font style='text-decoration:none'>"
		ijunta="<img src='../images/tjunta.jpg' width='16' height='14' border='0' alt=''>"
		isepara="<img src='../images/tsepara.jpg' width='16' height='14' border='0' alt=''>"
		if rs2.eof then 
			texto=texto & "<div align="center" id='0' alt='1' onclick=""NewWindow('gradenovaaula.asp?idaula=0&idturma=" & id_grdturma & "&codhor=" & codhor & "','InclusaoGradev2','650','400','yes','center');return false"" onfocus='this.blur()'><hr>Aula Vaga</div>"
			classe="campoa"
		end if
	'<img src="../images/espelho.jpg" width="16" height="16" border="0" alt="">
		do while not rs2.eof
			imagem="":nome1="":nome2=""
			if rs2("juntar")=true then imagem=imagem&"<br><font color=red><b>-><-" & rs2("jturma")& "</b></font>":imagem2="-><-" & rs2("jturma")
			if rs2("dividir")=true then imagem=imagem&"<br><font color=red><b><-->" & rs2("dturma")& "</b></font>":imagem2="<-->" & rs2("dturma")
			if isnull(rs2("chapa1"))=false then
				sql8="select nome from grades_aux_prof2 where chapa='" & rs2("chapa1") & "' "
				rs3.Open sql8, ,adOpenStatic, adLockReadOnly
				if rs3.recordcount>0 then nome1=rs3("nome") else nome1="-"
				if rs2("chapa1")="99999" then nome1="SEM PROFESSOR"
				rs3.close
			end if
			if isnull(rs2("chapa2"))=false then
				sql8="select nome from grades_aux_prof2 where chapa='" & rs2("chapa2") & "' "
				rs3.Open sql8, ,adOpenStatic, adLockReadOnly
				if rs3.recordcount>0 then nome2=rs3("nome") else nome2="-"
				if rs2("chapa2")="99999" then nome2="SEM PROFESSOR"
				rs3.close
			end if
			'if periodolanc<now() and session("usuariomaster")<>"02379" then
			if (periodolanc<int(now())+1 or aberto=0) and session("usuariomaster")<>"02379" then
				texto=texto & "<div style='border:0px dashed' id='" & rs2("id_grdaula") & "' alt='1' >"
			else
				texto=texto & "<div style='border:0px dashed' id='" & rs2("id_grdaula") & "' alt='1' onclick=""NewWindow('gradenovaaula.asp?idaula=" & rs2("id_grdaula") & "&idturma=" & id_grdturma & "&codhor=" & codhor  & "','InclusaoGradev2','650','400','yes','center');return false"" onfocus='this.blur()'>"
			end if
			'if rs2.absoluteposition<registros then texto=texto & risc1 else texto=texto & risc2
			tam1=19:if len(rs2("materia"))>tam1 then materia=left(rs2("materia"),tam1) & "..." else materia=rs2("materia")
			tam2=20:if len(nome1)>tam2 then nome1=left(nome1,tam2) & "..."
			tam3=20:if len(nome2)>tam3 then nome2=left(nome2,tam3) & "..."
			if rs2.absoluteposition<registros and rs2("termino")<>terminoperiodo then 
				texto=texto & "<img src='../images/espelho.jpg' width='16' height='16' border='0' alt='"
				texto=texto & rs2("materia") & chr(10) & chr(13) & nome1 & chr(10) & chr(13)
				if nome2<>"" then texto=texto & nome2 & chr(10) & chr(13)
				texto=texto & rs2("inicio") & " a " & rs2("termino") & chr(10) & chr(13) & imagem2
				texto=texto & "'>" & "<font color=gray>" & rs2("inicio") & " a " & rs2("termino") & "</font>" '& imagem
			else
				if rs2("termino")<terminoperiodo then  texto=texto & risc1 else texto=texto & risc2
				texto=texto & "<font color=black><b>" & materia & "</b></font>" & "<br>"
				texto=texto & "<font color=blue>" & nome1 & "</font><br>"
				if nome2<>"" then texto=texto & "<font color=green>" & nome2 & "</font><br>"
				texto=texto & "<font color=gray>" & rs2("inicio") & " a " & rs2("termino") & "</font>" & imagem & "<br>"
				texto=texto & "</font>"
			end if
			texto=texto & "</div>"
			classe="campor"
		rs2.movenext:loop
		rs2.close
	else
		texto="<hr><div id='0' alt='0'>"
		classe="fundo"
	end if
%>
	<td class=<%=classe%> <%=borda%> valign=top nowrap width=150>
	<%=texto%>
	</td>
<%
next
%>
</tr>
<%
rs.movenext:loop
rs.close
response.write texto1
%>

</table>
<%
end if ' para tipocurso=2


'---------------******************-------------------******************-----------------------*********************
if tipocurso="4" or tipocurso="5" or tipocurso="6" then
%>
<%
dim emes(12),edia(220),fdia(40), mdia(220), ndia(220), iddia(220), iddata(220)
d_zero=iniciocalendar
mesagora=month(iniciocalendar)
anoagora=year(iniciocalendar)
diaagora=day(iniciocalendar)

sqld="select diaferiado as dia1 from corporerm.dbo.gferiado " & _
"where diaferiado between '" & dtaccess(inicioperiodo) & "' and '" & dtaccess(terminoperiodo) & "' "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof 
	dias=rs2("dia1")-d_zero
	edia(dias)=1
rs2.movenext:loop
end if
rs2.close

sqla="select a.id_grdturma, a.id_grdaula, a.codmat, m.materia, a.chapa1, f.nome, d.data, d.id_data " & _	
"from g2aulas a, (select id_grdaula, data, id_data from g2aulasdata where deletada=0 group by id_Grdaula, data, id_data) d, grades_aux_prof f, corporerm.dbo.umaterias m " & _
"where a.id_grdaula=d.id_grdaula and f.chapa=a.chapa1 and a.codmat=m.codmat collate database_default " & _
"and a.id_grdturma=" & id_grdturma & " and d.data between '" & dtaccess(inicioperiodo) & "' and '" & dtaccess(terminoperiodo) & "' " & _
"group by a.id_grdturma, a.id_grdaula, a.codmat, m.materia, a.chapa1, f.nome, d.data, d.id_data " 
rs2.Open sqla, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof 
	dias=rs2("data")-d_zero
	'tdia(dias)=1
	if mdia(dias)<>"" then mdia(dias)="(2) " & mdia(dias) & "<br>" & rs2("materia") else mdia(dias)=rs2("materia")
	if ndia(dias)<>"" then ndia(dias)="(2) " & ndia(dias) & "<br>" & rs2("nome") else ndia(dias)=rs2("nome")
	'mdia(dias)=mdia(dias) & rs2("materia")
	'ndia(dias)=ndia(dias) & rs2("nome")
	iddia(dias)=rs2("id_grdaula")
	iddata(dias)=rs2("id_data")
rs2.movenext:loop
end if
rs2.close
%>

<!-- calencario -->
<%
tamlinha=45:tamcol=50
emes(1)="Janeiro":emes(2)="Fevereiro":emes(3)="Março":emes(4)="Abril":emes(5)="Maio":emes(6)="Junho"
emes(7)="Julho":emes(8)="Agosto":emes(9)="Setembro":emes(10)="Outubro":emes(11)="Novembro":emes(12)="Dezembro"

if month(terminoperiodo)<month(iniciocalendar) then termino=12 else termino=month(terminoperiodo)
for zcal=month(iniciocalendar) to termino
if zcal=month(iniciocalendar) then response.write "<table border='0' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse'><tr><td valign='top'>"
if coluna<3 and coluna>0 then
	response.write "</td><td valign='top'>"
else
	response.write "</td></tr><tr><td valign='top'>":coluna=0
end if
mesagora=zcal
anoagora=year(iniciocalendar)
diaagora=1
diasemana=weekday(dateserial(anoagora,mesagora,1))
ultimodia=day(dateserial(anoagora,mesagora+1,1)-1)
ultimo=0
%>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=tamcol*7%>">
<tr>
	<td class=fundo width="100%" align="center">
		<font color="#000080"><b><%=emes(mesagora)& "/" & anoagora%></font></td>
</tr>
</table>
<table border="1" bordercolor="#cccccc" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="<%=tamcol*7%>">
<tr>
	<td class="campo" align="center" style="color:red">Dom</td>
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
		ultimo=1:if fdia(ultimo)=1 then fundo="fundor" else fundo="campor"
		testadata=dateserial(anoagora,mesagora,ultimo)-d_zero
		if edia(testadata)=1 then 'é feriado - 'edia(ultimo) passa a edia(testadata) por ser multiplos meses
			response.write "<td class=" & fundo & " align='left' valign='top' height='" & tamlinha & "' width='" & tamcol & "'>"
			'response.write "<a href='#' class='r'>"
			'if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write "<font style=color:red>" & ultimo
			'response.write "</a>"
			buscadb=0
		else
			response.write "<td class=" & fundo & " align='left' valign='top' height='" & tamlinha & "' width='" & tamcol & "'>"
			'response.write "<a href='#' class='r'>"
			'if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			'response.write "</a>"
			buscadb=1
		end if		
	elseif ultimo>=1 then
		ultimo=ultimo+1:if fdia(ultimo)=1 then fundo="fundor" else fundo="campor"
		testadata=dateserial(anoagora,mesagora,ultimo)-d_zero
		if edia(testadata)=1 then 'edia(ultimo) passa a edia(testadata) por ser multiplos meses
			response.write "<td class=" & fundo & " align='left' valign='top' height='" & tamlinha & "' width='" & tamcol & "'>"
			'response.write "<a href='#' class='r'>"
			'if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write "<font style=color:red>" & ultimo
			'response.write "</a>"
			buscadb=0
		else
			response.write "<td class=" & fundo & " align='left' valign='top' height='" & tamlinha & "' width='" & tamcol & "'>"
			'response.write "<a href='#' class='r'>"
			'if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			'response.write "</a>"
			buscadb=1
		end if
	else
		response.write "<td class="campor" align='center'>"
		buscadb=0
	end if
'-----------------------------
if buscadb=1 and (mdia(testadata)<>"") then
	x=mdia(testadata) & "<br>" & ndia(testadata)
	response.write "<div style='background-image:url();border:0px dashed' id='" & iddia(testadata) & "' onclick=""NewWindow('gradenovaposdata.asp?id_grdaula=" & id_grdaula & "&id_data=" & iddata(testadata)  & "','InclusaoGradev2','545','250','yes','center');return false"" onfocus='this.blur()' onMouseover=""ddrivetip('" & x & "', 250)""; onMouseout='hideddrivetip()'>"		
	response.write "<font style='font-size:7pt'>" & left(mdia(testadata),7)
	response.write "<br>"
	response.write "<font style='font-size:6pt'>" & left(ndia(testadata),9)
	response.write "</div>"
end if
'-----------------------------
	response.write "</td>"
next
response.write "</tr>"

vartemp1=ultimodia-ultimo
vartemp2=int(vartemp1/7)
if (vartemp1/7)-vartemp2>0 then vartemp2=vartemp2+1
for sem=1 to vartemp2
	response.write "<tr>"
	for l2=1 to 7
		ultimo=ultimo+1:if fdia(ultimo)=1 then fundo="fundor" else fundo="campor"
		testadata=dateserial(anoagora,mesagora,ultimo)-d_zero
		response.write "<td class=" & fundo & " align='left' valign='top' height='" & tamlinha & "' width='" & tamcol & "'>"
		if ultimo<=ultimodia then 
			if edia(testadata)=1 or l2=1 then 'edia(ultimo) passa a edia(testadata) por ser multiplos meses
				'response.write "<a href='#'>"
				'if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write "<font style=color:red>" & ultimo
				'response.write "</a>"
				buscadb=0
			else
				'response.write "<a href='#'>"
				'if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				'response.write "</a>"
				buscadb=1
if buscadb=1 and (mdia(testadata)<>"") then
	x=mdia(testadata) & "<br>" & ndia(testadata)
	response.write "<div style='background-image:url();border:0px dashed' id='" & iddia(testadata) & "' onclick=""NewWindow('gradenovaposdata.asp?id_grdaula=" & iddia(testadata) & "&id_data=" & iddata(testadata)  & "','InclusaoGradev2','545','250','yes','center');return false"" onfocus='this.blur()' onMouseover=""ddrivetip('" & x & "', 250)""; onMouseout='hideddrivetip()'>"		
	response.write "<font style='font-size:7pt'>" & left(mdia(testadata),7)
	response.write "<br>"
	response.write "<font style='font-size:6pt'>" & left(ndia(testadata),9)
	response.write "</div>"
end if
			end if
		end if
'-----------------------------
'-----------------------------
		response.write "</td>"
	next
	response.write "</tr>"
next
%>
</table>
<!-- fim calencario -->

<%
coluna=coluna+1
next 'zcal
response.write "</tr></table>"
%>



<%
end if ' para tipocurso=4,5,6
%>

<!-- -->
</td></tr><tr><td valign=top style="border-right:2px solid black;border-top:2px solid black">
<!-- -->

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" >
<tr>
<td class=titulo height=15 colspan=6>Disciplina da grade</td></tr>
<tr>
	<td class="campop">
	<input type="radio" name="seldisc" value="0" <%if request.form("seldisc")="0" then response.write "checked"%> onclick="javascript:submit()" >
	</td>
	<td class="campol" align="center">Matéria</td>
	<td class="campol" align="center">Aulas<br>Semanais</td>
	<td class="campol" align="center">Carga<br>Horária</td>
	<td class="campol" align="center">Aulas<br>atribuidas</td>
	<td class="campop"></td>
</tr>
<%
if session("usuariomaster")="02379" then response.write terminoperiodo & "<---"
datagrade=now()
if now()>terminoperiodo then datagrade=terminoperiodo 'else datagrade=now()
if now()<inicioperiodo then datagrade=inicioperiodo
if session("usuariomaster")="02379" then response.write datagrade & "<---"
if tipocurso="2" then sql5="SELECT g.CODMAT, m.MATERIA, g.NAULASSEM, g.CARGAHORARIA, z.total " & _
"FROM (corporerm.dbo.ugrade g INNER JOIN corporerm.dbo.umaterias m ON g.codmat=m.codmat) " & _
"LEFT JOIN (select codmat, count(codmat) total from g2aulas where id_grdturma=" & id_grdturma & " and deletada=0 and ativo=1 and '" & dtaccess(DataGrade) & "' between inicio and termino group by codmat) z ON g.codmat=z.codmat collate database_default " & _
"where codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and periodo=" & serie & _
" order by m.materia"
if tipocurso<>"2" then sql5="SELECT g.CODMAT, m.MATERIA, g.NAULASSEM, g.CARGAHORARIA, z.total " & _
"FROM (corporerm.dbo.ugrade g INNER JOIN corporerm.dbo.umaterias m ON g.codmat=m.codmat) " & _
"LEFT JOIN (select codmat, count(h.codhor) total from g2aulas a, g2aulasdata d, g2aulashora h where id_grdturma=" & id_grdturma & " and d.id_grdaula=a.id_grdaula and h.id_data=d.id_data and a.deletada=0 and d.deletada=0 and ativo=1 group by codmat) z ON g.codmat=z.codmat collate database_default " & _
"where codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and periodo=" & serie & _
" order by m.materia"
if tipocurso="2" then sql5="SELECT g.CODMAT, m.MATERIA, g.NAULASSEM, g.CARGAHORARIA, z.total " & _
"FROM (corporerm.dbo.ugrade g INNER JOIN corporerm.dbo.umaterias m ON g.codmat=m.codmat) " & _
"LEFT JOIN (select codmat, count(codmat) total from g2aulas where id_grdturma=" & id_grdturma & " and deletada=0 and ativo=1 and '" & dtaccess(DataGrade) & "' between inicio and termino group by codmat) z ON g.codmat=z.codmat collate database_default " & _
"where codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and periodo=" & serie & _
" order by m.materia"
if tipocurso<>"2" then sql5="SELECT g.CODMAT, m.MATERIA, g.NAULASSEM, g.CARGAHORARIA, z.total " & _
"FROM (corporerm.dbo.ugrade g INNER JOIN corporerm.dbo.umaterias m ON g.codmat=m.codmat) " & _
"LEFT JOIN (select codmat, sum(qtaulas) total from g2aulas a, g2aulasdata d " & _
	"where id_grdturma=" & id_grdturma & " and d.id_grdaula=a.id_grdaula and a.deletada=0 " & _
	"and d.deletada=0 and ativo=1 group by codmat) z ON g.codmat=z.codmat collate database_default " & _
"where codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and periodo=" & serie & _
" order by m.materia"

rs2.Open sql5, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
if rs2("total")="" or isnull(rs2("total")) then teste1=0 else teste1=cdbl(rs2("total"))
if tipocurso="2" then 
	if rs2("naulassem")="" or isnull(rs2("naulassem")) then teste2=0 else teste2=cdbl(rs2("naulassem"))
else
	if rs2("cargahoraria")="" or isnull(rs2("cargahoraria")) then teste2=0 else teste2=cdbl(rs2("cargahoraria"))
end if
if teste1<>teste2 then corfont4="<font color=red>" else corfont4="<font color=black>"
%>
<tr>
	<td class="campop">
	<%if tipocurso<>"2" then %><a href="gradenovapos.asp?idturma=<%=id_grdturma%>&codmat=<%=rs2("codmat")%>" onclick="NewWindow(this.href,'InclusaoGradev2pos','650','600','yes','center');return false" onfocus="this.blur()">
	<img src="../images/mais.gif" width="10" height="10" border="0" alt=""></a>
	<%end if%>	
	<%if tipocurso="2" then %>
	<input type="radio" name="seldisc" value="<%=rs2("codmat")%>/<%=request.form("selprof"&rs2.absoluteposition)%>" <%if request.form("seldisc")=rs2("codmat")&"/"&request.form("selprof"&rs2.absoluteposition) then response.write "checked"%> onclick="javascript:submit()" >
	<%end if%>	
	</td>
	<td class="campolr"><div name="<%=rs2("materia")%>"><b><%=rs2("materia")%></b> (<font color=gray><%=rs2("codmat")%></font>)</div>
	<%
	achou=0
	sqlprof="select distinct chapa1, nome from g2ch g inner join grades_aux_prof2 f on f.chapa=g.chapa1 " & _
	"where codmat='" & rs2("codmat") & "' and deletada=0 and termino='" & terminoanterior & "' " & _
	"and left(codtur,7)=left('" & codtur & "',7) and coddoc='" & coddoc & "' and codsituacao in ('A','F','Z','E')"
	rs3.Open sqlprof, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount=1 then 'achou professor do semestre anterior
	achou=1
	chapadisc=rs3("chapa1")
	chapanome=rs3("nome")
	%>
	<input type="hidden" name="selprof<%=rs2.absoluteposition%>" value="<%=chapadisc%>">
	<%
	response.write "<font color='blue'>" & chapanome & "</font>"
		sqllimite="select chapa, limite from g2limite where chapa='" & chapadisc & "'"
		rs.Open sqllimite, ,adOpenStatic, adLockReadOnly
		if rs.recordcount=1 then limite=rs("limite") else limite=20
		rs.close
		sqlatrib="select count(chapa1) as taulas from g2ch where chapa1='" & chapadisc & "' and '" & dtaccess(datagrade) & "' between inicio and termino and juntar=0 "
		rs.Open sqlatrib, ,adOpenStatic, adLockReadOnly
		taulas1=rs("taulas")
		rs.close
	end if
	rs3.close
	if achou=0 then ' (A) ainda não econtrou professor na disciplina/turma no semestre anterior
	sqlprof="select distinct chapa1, nome from g2ch g inner join grades_aux_prof2 f on f.chapa=g.chapa1 " & _
	"where codmat='" & rs2("codmat") & "' and deletada=0 and coddoc='" & coddoc & "' and codsituacao in ('A','F','Z','E') "
	rs3.Open sqlprof, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then 'achou em semestres anteriores
	achou=1
	%>
	<select size=1 name="selprof<%=rs2.absoluteposition%>" onchange="javascript:submit()">
	<%
	do while not rs3.eof
		if request.form("selprof" & rs2.absoluteposition)=rs3("chapa1") then txtsel1="selected" else txtsel1=""
	%>
		<option <%=txtsel1%> value="<%=rs3("chapa1")%>"><%=rs3("nome")%></option>
	<%
	rs3.movenext:loop
	%>
	</select>
	<%
	
	end if ' recordset3>0
	rs3.close

	end if ' achou=0 (A)


	response.write "<font color=#FF0000> Limite: " & limite & "</font>"
	response.write "<font color=#9900FF> Atribuidas: " & taulas1 & "</font>"
		
	%>
	</td>
	<td class="campol" align="center"><%=rs2("naulassem")%></td>
	<td class="campol" align="center"><%=rs2("cargahoraria")%></td>
	<td class="campol" align="center" style="border-right:2 solid 000000"><%=corfont4%><%=rs2("total")%></td>
	<td class="campor">
<%
sql6="select status, descricao, count(mataluno) alunos from corporerm.dbo.umatalun m, corporerm.dbo.usitmat s where s.codsitmat=m.status and " & _
"perletivo='" & perlet & "' and codmat='" & rs2("codmat") & "' and codtur='" & codtur & "' group by status, descricao "
rs3.Open sql6, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
do while not rs3.eof 
	
	if rs3("descricao")="CURSANDO" or rs3("descricao")="APROVADO" then 
		corfonte1="<font color=blue><b>" : final3="</b>" : corfonte2="<font color=red>"
		response.write corfonte1 & rs3("descricao") & ": " & corfonte2 & rs3("alunos") & final3
	end if
	corfonte1="<font color=black>":corfonte2="<font color=black>":final3=""
	if session("usuariomaster")="02379" and (rs3("descricao")<>"CURSANDO" and rs3("descricao")<>"APROVADO") then
		response.write corfonte1 & rs3("descricao") & ": " & rs3("alunos")
	end if
	if rs3.absoluteposition<rs3.recordcount then response.write " / "
rs3.movenext:loop
end if
rs3.close
%>
	</td>
</tr>
<%
rs2.movenext:loop
rs2.close
%>
<tr>
	<td class=campo colspan=6>
<%
'if session("usuariomaster")="02379" or session("usuariogrupo")="RH" then
sqlm="select m.status, s.descricao, count(mataluno) alunos from corporerm.dbo.umatricpl m, corporerm.dbo.usitmat s where m.status=s.codsitmat and status not in ('40','53','92') " & _
"and m.perletivo='" & perlet & "' and codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and periodo=" & serie & " " & _
"group by m.status, s.descricao "
rs2.Open sqlm, ,adOpenStatic, adLockReadOnly
response.write "<font color=red>Informação adicional para a/o " & serie & " série/ano (todos os periodos: mat/not/int):<br>"
do while not rs2.eof
	response.write "<font color=black>"
	response.write rs2("status")
	response.write " - "
	response.write rs2("descricao")
	response.write " -> "
	response.write "<font color=blue>"
	response.write rs2("alunos")
	total1=total1+rs2("alunos")
	if rs2.absoluteposition=rs2.recordcount then response.write "<font color=green> = " & total1 else response.write "<br>"
rs2.movenext
loop
rs2.close
'end if
%>
</td>
</tr>
</table>
<br>
<%

sql="select initurma=min(serie), fimturma=max(serie) from g2turmas where turmaid='" & turmaid & "' and codtur<>'" & codtur & "' and serie<" & serie
rs2.Open sql, ,adOpenStatic, adLockReadOnly
initurma=rs2("initurma"):fimturma=rs2("fimturma")
rs2.close
if isnull(initurma) then initurma=1
if isnull(fimturma) then fimturma=1
if turmaid<>"" then 

for a=initurma to fimturma
	sql2="select distinct serie, perlet from g2turmas where turmaid='" & turmaid & "' order by serie"
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
	redim preserve hperlet(rs2("serie"))
	hperlet(rs2("serie"))=rs2("perlet")
	rs2.movenext
	loop
	rs2.close
next
for a=initurma to fimturma
	sql2="select serie, codtur from g2turmas where turmaid='" & turmaid & "' order by serie, codtur "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
	redim preserve hcodtur(rs2("serie"))
	redim preserve hcodtur2(rs2("serie"))
	if hcodtur(rs2("serie"))<>"" and rs2.absoluteposition>1 then
		hcodtur2(rs2("serie"))=rs2("codtur")
	else
		hcodtur(rs2("serie"))=rs2("codtur")
	end if
	rs2.movenext
	loop
	rs2.close
next
for a=initurma to fimturma
	sql2="select serie, codtur, codcur, codper, grade from g2turmas where turmaid='" & turmaid & "' order by serie, codtur "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
	redim preserve hcodtur(rs2("serie"))
	redim preserve hcodtur2(rs2("serie"))
	redim preserve hcodcur(rs2("serie"))
	redim preserve hcodper(rs2("serie"))
	redim preserve hgrade(rs2("serie"))
	if hcodtur(rs2("serie"))<>"" and rs2.absoluteposition>1 then
		hcodtur2(rs2("serie"))=rs2("codtur")
	else
		hcodtur(rs2("serie"))=rs2("codtur")
	end if
	hcodcur(rs2("serie"))=rs2("codcur")
	hcodper(rs2("serie"))=rs2("codper")
	hgrade(rs2("serie"))=rs2("grade")
	rs2.movenext
	loop
	rs2.close
next

'for a=initurma to fimturma
'	response.write hperlet(a) & "-" & hcodtur(a) & "-" & hcodtur2(a) &  "<br>"
'next

%>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" >
<tr>
	<td class=fundop colspan="<%=fimturma+1%>" align="center"><b>Evolução das turmas</b></td>
</tr>
</tr>
<tr>
	<td class=fundo>Série</td>
<%for a=initurma to fimturma%>
	<td class=campo align="center"><%=a%></td>
<%next%>
</tr>
<tr>
	<td class=fundo>Per.Letivo</td>
<%for a=initurma to fimturma%>
	<td class=campo align="center"><%=hperlet(a)%></td>
<%next%>
</tr>
<tr>
	<td class=fundo>Turma</td>
<%for a=initurma to fimturma%>
	<td class=campo align="center"><%=hcodtur(a)%></td>
<%next%>
</tr>
<tr>
	<td class=fundo></td>
<%for a=initurma to fimturma%>
	<td class=campo align="left" valign=top>
	<%
	sql3="select codcur, codper, grade, status, descricao, total=count(mataluno) " & _
	"from corporerm.dbo.umatricpl m inner join corporerm.dbo.usitmat s on s.codsitmat=m.status " & _
	"where codcur=" & hcodcur(a) & " and codper=" & hcodper(a) & " and grade=" & hgrade(a) & " and periodo=" & a & " and perletivo='" & hperlet(a) & "' " & _
	"group by codcur, codper, grade, status, descricao"
	sql3="select m.codcur, m.codper, m.grade, status, descricao, total=count(m.mataluno) " & _
	"from corporerm.dbo.umatricpl m inner join corporerm.dbo.usitmat s on s.codsitmat=m.status " & _
	"inner join (select distinct perletivo, mataluno, codcur, codper, grade, codtur from corporerm.dbo.umatalun " & _
	"where codtur='" & hcodtur(a) & "') a on a.mataluno=m.mataluno and a.codcur=m.codcur and a.codper=m.codper and a.grade=m.grade and a.perletivo=m.perletivo " & _
	"where m.codcur=" & hcodcur(a) & " and m.codper=" & hcodper(a) & " and m.grade=" & hgrade(a) & " and periodo=" & a & " and m.perletivo='" & hperlet(a) & "' " & _
	"group by m.codcur, m.codper, m.grade, status, descricao"
	rs2.Open sql3, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
		if len(rs2("descricao"))>12 then fonte="<font style='font-size:8px'>" else fonte="<font style='font-size:11px'>"
		response.write fonte & rs2("descricao") '& " ("  & rs2("status") & ")" 
		response.write ": "
		response.write "<font style='font-size:11px'>" & rs2("total")	
	rs2.movenext
	response.write "<br>"
	loop
	rs2.close
	%>
	</td>
<%next%>
</tr>

</table>
<%
end if ' turmaid<>""

 %>

<!-- -->
</td><td valign=top style="border-right:2px solid black;border-top:2px solid black">
<!-- -->
<%

if tipocurso="2" or tipocurso="12" then

sqlcusto="select custo=sum(custo), adic=sum(adicionais), enc=sum(enc_mes), prov=sum(provisao) " & _
"from ( " & _
"select g.chapa1, g.coddoc, ta, valoraula=(case when valoraula is null then 29.02 else valoraula end), " & _
"custo=round(ta*4.5*(case when valoraula is null then 29.02 else valoraula end),2), " & _
"adicionais=round(ta*4.5*(case when valoraula is null then 29.02 else valoraula end)*0.225,2), " & _
"enc_mes=round(ta*4.5*(case when valoraula is null then 29.02 else valoraula end)*1.225*0.09,2), " & _
"provisao=round(ta*4.5*(case when valoraula is null then 29.02 else valoraula end)*1.225*0.2119,2) " & _
"from g2ch g " & _
"left join dc_professor f on f.chapa=g.chapa1 " & _
"left join salarios_curso_faixa s  on f.tab_instr=s.titulacao and f.codnivelsal=s.nivel and f.tab_ref=s.reformulacao and s.coddoc=g.coddoc " & _
"where g.id_grdturma=" & id_grdturma & " and g.deletada=0 and '" & dtaccess(inicioperiodo) & "' between inicio and termino " & _
") c "
'if session("usuariomaster")="02379" then response.write sqlcusto
rs2.Open sqlcusto, ,adOpenStatic, adLockReadOnly
if semturma=1 then
	custo=0: adic=0: enc=0: prov=0: custototal=0
else
	if isnull(rs2("custo")) then custo=0 else custo=rs2("custo")
	if isnull(rs2("adic"))  then adic =0 else adic =rs2("adic")
	if isnull(rs2("enc"))   then enc  =0 else enc  =rs2("enc")
	if isnull(rs2("prov"))  then prov =0 else prov =rs2("prov")
	custototal=custo+adic+enc+prov
end if
rs2.close
%>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" >
<tr>
	<td class=fundo style="border:1px solid" colspan=4 align="center"><b>Custo da Turma</b> (em fase de testes)
		<br><span style="font-size:9px">(não inclue coordenação, estágios, orientações)
		<br><span style="font-size:9px">()</td>
</tr>
<tr><td class=campo colspan=2 style="border:1px solid">Custo das aulas <span style="font-size:8px;vertical-align:top">(1)</span></td>
	<td class=campo colspan=2 style="border:1px solid" align="right"><%=formatnumber(custo,2)%></td></tr>
<tr><td class=campo colspan=2 style="border:1px solid">Adicionais (DSR/Hora atividade) <span style="font-size:8px;vertical-align:top">(2)</span></td>
	<td class=campo colspan=2 style="border:1px solid" align="right"><%=formatnumber(adic,2)%></td></tr>
<tr><td class=campo colspan=2 style="border:1px solid">Encargos e Provisões <span style="font-size:8px;vertical-align:top">(3)</span></td>
	<td class=campo colspan=2 style="border:1px solid" align="right"><%=formatnumber(enc+prov,2)%></td></tr>
<tr><td class=campo colspan=2 style="border:1px solid">Custo Estimado (mês)</td>
	<td class=campo colspan=2 style="border:1px solid" align="right"><%=formatnumber(custototal,2)%></td></tr>
<tr><td class=campo colspan=4 height=15></td></tr>
<%
sqle="select mens_m, mens_n, eva_media, eva_aberto, eva_m, eva_n from g2cursoeve where coddoc='" & coddoc & "'"
rs2.Open sqle, ,adOpenStatic, adLockReadOnly
if semturma=1 then
	mens_m=0:mens_n=0:eva_media=0:eva_aberto=0:eva_m=0:eva_n=0
else
	mens_m=rs2("mens_m"):mens_n=rs2("mens_n"):eva_media=rs2("eva_media"):eva_aberto=rs2("eva_aberto"):eva_m=rs2("eva_m"):eva_n=rs2("eva_n")
end if
%>
<tr><td class=campo colspan=4 style="border:1px solid" align="center"><b>Evasão (%)</td></tr>
<tr><td class=fundo style="border:1px solid" align="center">Média geral <span style="font-size:8px;vertical-align:top">(4)</span></td>
	<td class=fundo style="border:1px solid" align="center">P.Let.Atuais <span style="font-size:8px;vertical-align:top">(5)</span></td>
	<td class=fundo style="border:1px solid" align="center">Per.Mat. <span style="font-size:8px;vertical-align:top">(6)</span></td>
	<td class=fundo style="border:1px solid" align="center">Per.Not. <span style="font-size:8px;vertical-align:top">(7)</span></td>
</tr>
<tr><td class=campo style="border:1px solid" align="center"><%=eva_media%></td>
	<td class=campo style="border:1px solid" align="center"><%=eva_aberto%></td>
	<td class=campo style="border:1px solid" align="center"><%=eva_m%></td>
	<td class=campo style="border:1px solid" align="center"><%=eva_n%></td>
</tr>
<tr><td class=campo colspan=4 height=15></td></tr>
<%
if turno=3 then mensalidade=mens_n else mensalidade=mens_m
if eva_media=0 then evasao=eva_aberto else evasao=eva_media
if turno=1 and eva_m>0 then evasao=eva_m
if custototal>0 then
	pef=int(custototal/mensalidade)+1
	pei=int(pef/((100-evasao)/100))+1
else
	pef=0:pei=0
end if
%>
<tr>
	<td class=campo colspan=3 style="border:1px solid">Ponto de equilibrio inicial (1º semestre) <span style="font-size:8px;vertical-align:top">(8)</span></td>
	<td class=campo colspan=1 style="border:1px solid" align="center"><%=pei%></td>
</tr>
<tr>
	<td class=campo colspan=3 style="border:1px solid">Ponto de equilibrio final (último semestre) <span style="font-size:8px;vertical-align:top">(9)</span></td>
	<td class=campo colspan=1 style="border:1px solid" align="center"><%=pef%></td>
</tr>
<tr><td class=campo colspan=4 height=5></td></tr>
<tr>
	<td class=campo colspan=4>
	<p style="font-size:10px;margin-bottom:0px;margin-top:0px">Notas:
	<br>(1) Custo das aulas semanais de cada professor x 4,5 semanas
	<br>(2) Adicionais legais de 5% hora atividade e 1/6 D.S.R.
	<br>(3) Fundo de Garantia e Pis sobre o valor das aulas e <br>reserva mensal para férias e 13º Salário (1/12)
	<br>(4) Comparação de todas as turmas concluidas entre o número<br>de alunos que estavam matriculados no primeiro semestre<br>
	e o número de alunos matriculados no último semestre.
	<br>(5) Média da evasão das turmas em andamento comparando o primeiro<br>semestre com o semestre atual de cada turma.
	<br>(6) Cálculo da evasão média das turmas concluidas do periodo matutino.
	<br>(7) Cálculo da evasão média das turmas concluidas do periodo noturno.
	<br>(8) Ponto de equilíbrio considerando a mensalidade de <%=formatnumber(mensalidade,2)%><br> em relação ao custo da turma acrescida da taxa de evasão média.
	<br>(9) Ponto de equilíbrio considerando a mensalidade de <%=formatnumber(mensalidade,2)%><br> em relação ao custo da turma.
</td>
</tr>
	
</table>
<%
end if 'tipo 2 para custo.
%>

<!-- -->
</td></tr></table>
<!-- -->
<%
termino=now()
duracao=termino-inicio
response.write "<font size=1>" & formatdatetime(now()-inicio,3)
%>
</form>

</body>
</html>
<%
'rs.close
set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
set conexao=nothing
%>