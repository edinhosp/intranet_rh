<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a91")="N" or session("a91")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
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
<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rsc=server.createobject ("ADODB.Recordset")
set rsc.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("secao")="" then session("sel91")="Todas" else session("sel91")=request.form("secao")
	if session("sel91")<>"Todas" then
		'session("sqlb91")="AND (g.curso='" & session("sel91") & "') "
		session("sqlb91")="AND (g.coddoc='" & session("sel91") & "') "
	else
		session("sqlb91")=""
	end if

	if request.form("periodo")="" then session("per91")="Todos" else session("per91")=request.form("periodo")
	if session("per91")<>"Todos" then
		session("sqlc91")="AND (g.perlet='" & session("per91") & "') "
	else
		session("sqlc91")=""
	end if

	if request.form("perlanc")="" then session("lanc91")="Todos" else session("lanc91")=request.form("perlanc")
	if session("lanc91")<>"Todos" then
		session("sqlf91")="AND (g.perlet2 like '" & session("lanc91") & "') "
	else
		session("sqlf91")=""
	end if

	if request.form("turma")="" then session("tur91")="Todas" else session("tur91")=request.form("turma")
	if session("tur91")<>"Todas" then
		qserie=left(session("tur91"),len(session("tur91"))-1)
		qturma=right(session("tur91"),1)
		session("sqle91")="AND (g.serie=" & qserie & " and g.turma='" & qturma & "') "
	else
		session("sqle91")=""
	end if

	if request.form("diasem")="" then session("dia91")="Todos" else session("dia91")=request.form("diasem")
	if session("dia91")<>"Todos" then
		session("sqlg91")="AND (g.diasem=" & session("dia91") & ") "
	else
		session("sqlg91")=""
	end if
		
	if request.form("localizar")="" then
		session("loc91")=""
	else
		session("loc91")=request.form("localizar")
	end if
	if isnumeric(session("loc91"))=true then session("loc91")=numzero(session("loc91"),5)
	if session("loc91")<>"" then
		if isnumeric(session("loc91")) then
			session("sqld91")="AND (g.chapa1 like '%" & session("loc91") & "%' or g.chapa2 like '%" & session("loc91") & "%' ) "
		else
			session("sqld91")="AND (f.nome like '%" & session("loc91") & "%') "
		end if
	else
		session("sqld91")=""
	end if

		
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
if request.form("vez1")<>"Não" then 
	session("sqlc91")=session("sqlc91")
	session("per91")=session("per91")
else
	'session("sqlc91")="AND (g.perlet like '2005%') "
	'session("per91")="2005"
	session("sqlc91")=session("sqlc91")
	session("per91")=session("per91")
end if
registros=Session("RegistrosPorPagina")

sqla="SELECT g.*, f.nome AS PROFESSOR, gp.pini, gp.pfim, gp.lanc " & _
"FROM ((grades_5 AS g LEFT JOIN grades_aux_prof AS f ON g.chapa1=f.chapa) " & _
"INNER JOIN grades_aux_per AS gp ON (g.enfase=gp.enfase) AND (g.coddoc=gp.coddoc) AND (g.perlet2=gp.perlet2) AND (g.perlet=gp.perlet)) " & _
"WHERE g.id_grade>0 AND g.deletada=0  "
'AND g.coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "')
sqlb=""
sqlc="order by g.perlet, g.perlet2, g.curso, g.serie, g.turno, g.coddoc, g.diasem, g.a5, g.a3, g.a1 "

sql1=sqla & sqlb & session("sqlb91") & session("sqlc91") & session("sqld91") & session("sqle91") & session("sqlf91") & session("sqlg91") & sqlc

if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	'conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez")="Nao"
else
	if request("folha")="" then pagina=1
	if request.form("pagina")<>"" then pagina=request.form("pagina")
	if request("folha")<>"" then pagina=request("folha")
	Session("Pagina")=pagina
	conexao.cursorlocation = 3 'aduseclient
	'conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 	MostraDados
end if	

Sub MostraDados()
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
End Sub
%>
<form method="POST" action="grades.asp" name="form">
<input type="hidden" name="vez1" value="<%=session("PrimeiraVez")%>">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Grade Horária</p>
<table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo width="60%" valign="center" align="left">Página: 
<%
Session("Load1")="1"
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""grades.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""grades.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
end if

response.write "&nbsp;<b>"
response.write "<select size='1' name='pagina' onChange='javascript:submit()'>"
for selpag=1 to rs.pagecount
	if selpag=atual then selpag1="selected" else selpag1=""
	response.write "<option value=" & selpag & " " & selpag1 & ">" & selpag & "</option>"
next
response.write "</select>"
response.write "/" & rs.pagecount & "</b>&nbsp;"

if atual=rs.pagecount or rs.pagecount=0 then
response.write "<img src='../images/setanext0.gif' border='0'>"
response.write "<img src='../images/setalast0.gif' border='0'>"
else
response.write "<a href=""grades.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""grades.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="20%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
	<td class=campo width="20%" valign="top" align="right">
<% if session("a91")="T" then %>
<a href="grade_nova.asp" onclick="NewWindow(this.href,'InclusaoGrade','550','425','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
<font size="1">inserir novo horário</font></a>
<% end if %>
	</td>
</tr>
</table>

<table border="1" cellspacing="0" cellpadding="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulor align="center">Curso     </td>
	<td class=titulor align="center">Per.      </td>
	<td class=titulor align="center">Turma     </td>
	<td class=titulor align="center">Dia       </td>
	<td class=titulor align="center" colspan=6>Horário   </td>
	<td class=titulor align="center">Disciplina/Professor</td>
	<td class=titulor align="center">Sala       </td>
	<td class=titulor align="center">&nbsp;    </td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
'if rs("horini")="" then horini="&nbsp;" else horini=formatdatetime(rs("horini"),4)
'if rs("horfim")="" then horfim="&nbsp;" else horfim=formatdatetime(rs("horfim"),4)
if rs("chapa2")<>"" then
    sql="select nome from grades_aux_prof where chapa='" & rs("chapa2") & "' "
    rsc.Open sql, ,adOpenStatic, adLockReadOnly
    if rsc.recordcount>0 then professor2=rsc("nome") else professor2=""
    rsc.close
end if
professor1=rs("professor")
msg=""
if rs("pini")<>rs("inicio") then msg="Iniciou: " & rs("inicio")
if rs("pfim")<>rs("termino") then msg="Encerrou: " & rs("termino")
if rs("turno")="1" then turno="Mat"
if rs("turno")="2" then turno="Vesp"
if rs("turno")="3" then turno="Not"
if rs("turno")="31" then turno="Not"
if rs("turno")="5" then turno="Vesp-EF"
if rs("turno")="61" then turno="Int-M"
if rs("turno")="62" then turno="Int-V"
if rs("turno")="51" then turno="Mat"
if rs("turno")="52" then turno="Vesp"
if rs("turno")="53" then turno="Not"
if mid(rs("perlet2"),5,1)="A" then
	estilocel="campolr":tipopl="&nbsp<i>(Anual)</i>"
else
	estilocel="campotr":tipopl="&nbsp<i>(Semestral)</i>"
end if
semestre=right(rs("perlet2"),1)
ano=left(rs("perlet2"),4)
titulo=semestre & "º semestre / " & ano
if titulo<>lasttitulo then
	response.write "<tr>"
	response.write "<td class=grupo colspan=13>" & titulo & "</td>"
	response.write "</tr>"
end if
'if lastdiasem<>rs("diasem") then formato="style='border-top: 2px solid #000000'" else formato=""
if rs("a1")=1 then a1="" else a1=""
%>
<tr>
	<td class=<%=estilocel%> <%=formato%> ><%=rs("curso") %><%=tipopl%><br><%=rs("perlet")%></td>
	<td class=<%=estilocel%> <%=formato%> align="center"><%=turno%></td>
	<td class=<%=estilocel%> <%=formato%> align="center" nowrap><%=rs("codtur")%></td>
	<td class=<%=estilocel%> <%=formato%> ><%=weekdayname(rs("diasem"),1) %></td>

	<td class=<%=estilocel%> <%=formato%> nowrap>
	<%if rs("a1")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class=<%=estilocel%> <%=formato%> nowrap>
	<%if rs("a2")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class=<%=estilocel%> <%=formato%> nowrap>
	<%if rs("a3")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class=<%=estilocel%> <%=formato%> nowrap>
	<%if rs("a4")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class=<%=estilocel%> <%=formato%> nowrap>
	<%if rs("a5")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class=<%=estilocel%> <%=formato%> nowrap>
	<%if rs("a6")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>

	<td class=<%=estilocel%> <%=formato%> ><%=rs("materia") %>&nbsp;<font color="#FF0000"><%=msg%></font><br><font color="#0000FF"><%=professor1%></font> (<%=rs("chapa1")%>)
	<%if rs("chapa2")<>"" then%>
	<br><font color="#0000FF"><%=professor2%></font> (<%=rs("chapa2")%>)
	<%if rs("autorizado")=0 then response.write " <font color=red>(Ainda s/autorização)"
	end if%>
	</td>
	<td class=<%=estilocel%> <%=formato%> >
	<%=rs("codsala")%>
	</td>
	<td class=<%=estilocel%> <%=formato%> >
	<% if session("a91")="T" then %>
	<% if rs("lanc")=-1 or session("usuariomaster")="02379" then %>
		<a href="grade_alteracao.asp?codigo=<%=rs("id_grade")%>" onclick="NewWindow(this.href,'AlteracaoGrade','550','425','no','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% else %>
		<img src="../images/stop.gif" border="0" alt="<%=rs("id_grade")%>"></a>
	<% end if %>

	<% end if %>
	</td>
</tr>
<%
lastdiasem=rs("diasem")
if linha=1 then linha=0 else linha=1
lasttitulo=semestre & "º semestre / " & ano
rs.movenext
if rs.eof then exit for
'loop
Next

else 'sem registros
%>
<td class=grupo colspan=13>Esta seleção não mostra nenhum registro.</td>
<%
end if
%>
</table>

<p><font size="1">
<%

datainicial=dtaccess(dateserial(year(now())-2,month(now()),day(now())))

sql2="SELECT gc.coddoc, gc.CURSO " & _
"FROM grades_5 gc " & _
"WHERE inicio>'" & datainicial & "' " & _
"GROUP BY gc.coddoc, gc.CURSO ORDER BY gc.CURSO; "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Curso: <select size="1" name="secao">
<option value="Todas" <%if session("sel91")="Todas" then response.write "selected"%>>Todos cursos</option>
<%
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("coddoc")%>" <%if session("sel91")=rs2("coddoc") then response.write "selected"%>><%=rs2("curso")%> (<%=rs2("coddoc")%>)</option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<!-- turmas -->
&nbsp;&nbsp;Turma: <select size="1" name="turma">
<option value="Todas" <%if session("tur91")="Todas" then response.write "selected"%>>Todas</option>
<%
sql2="SELECT serie, turma FROM grades_5 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') and inicio>'" & datainicial & "' GROUP BY serie, turma "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
expr1=rs2("serie") & rs2("turma")
%>
	<option value="<%=expr1%>" <%if session("tur91")=expr1 then response.write "selected"%>><%=expr1%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<!-- Periodo letivo -->
&nbsp;&nbsp;Período Letivo: <select size="1" name="periodo" onchange="periodo1()">
<option value="Todos" <%if session("per91")="Todos" then response.write "selected"%>>Todos</option>
<%
sql2="SELECT perlet FROM grades_5 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') and inicio>'" & datainicial & "' GROUP BY perlet "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("perlet")%>" <%if session("per91")=rs2("perlet") then response.write "selected"%>><%=rs2("perlet")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<!-- Periodo lançamentos -->
&nbsp;&nbsp;Lançamentos: <select size="1" name="perlanc" onchange="perlanc1()">
<option value="Todos" <%if session("lanc91")="Todos" then response.write "selected"%>>Todos</option>
<%
'sql2="SELECT perlet2, right(perlet2,1) & 'º/' & left(perlet2,4) & ' ' & mid(perlet2,5,1) as descricao FROM grades_2 where codcur in (select codcur from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY perlet2, right(perlet2,1) & 'º/' & left(perlet2,4) & mid(perlet2,5,1) " & _
'"union all " & _
sql2="SELECT left(perlet2,4) + '%' + right(perlet2,1) as perlet2, right(perlet2,1) + 'º/' + left(perlet2,4) as descricao FROM grades_5 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') and inicio>'" & datainicial & "' GROUP BY left(perlet2,4) + '%' + right(perlet2,1), right(perlet2,1) + 'º/' + left(perlet2,4) "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("perlet2")%>" <%if session("lanc91")=rs2("perlet2") then response.write "selected"%>><%=rs2("descricao")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<br>
Localizar por nome/chapa: <input type="text" name="localizar" size=35 value="<%=session("loc91")%>">
&nbsp;&nbsp;Dia da semana: <select size="1" name="diasem" onchange="">
<option value="Todos" <%if session("dia91")="Todos" then response.write "selected"%>>Todos</option>
<%
for a=2 to 7
	extenso=weekdayname(a)
%>
	<option value="<%=a%>" <%if session("dia91")=cstr(a) then response.write "selected"%>><%=extenso%></option>
<%
next
%>
</select>

Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar">
</form>
</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>