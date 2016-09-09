<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a90")="N" or session("a90")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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
dim conexao, conexao2
dim rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rsc=server.createobject ("ADODB.Recordset")
set rsc.ActiveConnection = conexao

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("secao")="" then
		session("sel90")="Todas"
	else
		session("sel90")=request.form("secao")
	end if
	if session("sel90")<>"Todas" then
		session("sqlb90")="AND (g.coddoc='" & session("sel90") & "') "
	else
		session("sqlb90")=""
	end if

	if request.form("periodo")="" then
		session("per90")="Todos"
	else
		session("per90")=request.form("periodo")
	end if
	if session("per90")<>"Todos" then
		session("sqlc90")="AND (g.perlet like '" & session("per90") & "%') "
	else
		session("sqlc90")=""
	end if

	if request.form("perlanc")="" then
		session("lanc90")="Todos"
	else
		session("lanc90")=request.form("perlanc")
	end if
	if session("lanc90")<>"Todos" then
		session("sqlf90")="AND (g.perlet2 like '" & session("lanc90") & "') "
	else
		session("sqlf90")=""
	end if
		
	if request.form("localizar")="" then
		session("loc90")=""
	else
		session("loc90")=request.form("localizar")
	end if
	if isnumeric(session("loc90"))=true then session("loc90")=numzero(session("loc90"),5)
	if session("loc90")<>"" then
		if isnumeric(session("loc90")) then
			session("sqld90")="AND (g.chapa1 like '%" & session("loc90") & "%' or g.chapa2 like '%" & session("loc90") & "%' ) "
		else
			session("sqld90")="AND (f.nome like '%" & session("loc90") & "%') "
		end if
	else
		session("sqld90")=""
	end if

	if request.form("turma")="" then
		session("tur90")="Todas"
	else
		session("tur90")=request.form("turma")
	end if
	if session("tur90")<>"Todas" then
		qserie=left(session("tur90"),1)
		qturma=right(session("tur90"),1)
		session("sqle90")="AND (g.serie=" & qserie & " and g.turma='" & qturma & "') "
	else
		session("sqle90")=""
	end if

		
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
if request.form("vez1")<>"Não" then 
	session("sqlc90")=session("sqlc90")
	session("per90")=session("per90")
else
	session("sqlc90")="AND (g.perlet like '2005%') "
	session("per90")="2005"
end if
registros=Session("RegistrosPorPagina")

sqla="SELECT g.*, f.NOME AS PROFESSOR, gp.pini, gp.pfim, gp.lanc " & _
"FROM (grades_3 AS g LEFT JOIN grades_aux_prof AS f ON g.chapa1 = f.CHAPA) " & _
"INNER JOIN (select coddoc, perlet, perlet2, pini, pfim, lanc from grades_per where tper='L' group by coddoc, perlet, perlet2, pini, pfim, lanc ) AS gp ON (g.perlet=gp.perlet) AND (g.perlet2=gp.perlet2) AND (g.coddoc=gp.coddoc) " & _
"WHERE g.id_grade>0 AND g.deletada=0 and g.coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') "
sqlb=""
sqlc="order by g.perlet, g.perlet2, g.curso, g.serie, g.turma, g.diasem, g.turno, g.a5, g.a3, g.a1 "

sql1=sqla & sqlb & session("sqlb90") & session("sqlc90") & session("sqld90") & session("sqle90") & session("sqlf90") & sqlc

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
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Grade Horária - Cursos Livres</p>
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
<% if session("a90")="T" then %>
<a href="grade_nova.asp" onclick="NewWindow(this.href,'InclusaoGrade','535','320','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
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
if rs("turno")="71" then turno="Mat"
if rs("turno")="2" then turno="Vesp"
if rs("turno")="72" then turno="Vesp"
if rs("turno")="73" then turno="Not"
if rs("turno")="74" then turno="Not" 'alfa
if rs("turno")="75" then turno="Not"
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
	<td class=<%=estilocel%> <%=formato%> align="center"><%=rs("serie") & rs("turma") %></td>
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
	<%end if%>
	</td>
	<td class=<%=estilocel%> <%=formato%> >
	<%=rs("codsala")%>
	</td>
	<td class=<%=estilocel%> <%=formato%> >
	<% if session("a90")="T" then %>
	<% if rs("lanc")=-1 or session("usuariomaster")="02379" then %>
		<a href="grade_alteracao.asp?codigo=<%=rs("id_grade")%>" onclick="NewWindow(this.href,'AlteracaoGrade','550','320','no','center');return false" onfocus="this.blur()">
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
sql2="select curso, coddoc from grades_3 group by curso, coddoc having coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "')"
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Curso: <select size="1" name="secao">
<option value="Todas" <%if session("sel90")="Todas" then response.write "selected"%>>Todos cursos</option>
<%
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("coddoc")%>" <%if session("sel90")=rs2("coddoc") then response.write "selected"%>><%=rs2("curso")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<!-- turmas -->
&nbsp;&nbsp;Turma: <select size="1" name="turma">
<option value="Todas" <%if session("tur90")="Todas" then response.write "selected"%>>Todas</option>
<%
sql2="SELECT serie, turma FROM grades_3 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY serie, turma "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
expr1=rs2("serie") & rs2("turma")
%>
	<option value="<%=expr1%>" <%if session("tur90")=expr1 then response.write "selected"%>><%=expr1%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<!-- Periodo letivo -->
&nbsp;&nbsp;Período Letivo: <select size="1" name="periodo" onchange="periodo1()">
<option value="Todos" <%if session("per90")="Todos" then response.write "selected"%>>Todos</option>
<%
sql2="SELECT perlet FROM grades_3 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY perlet " & _
"union all " & _
"SELECT Left([perlet],4) AS Expr1 FROM grades_2 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY Left([perlet],4) "
sql2="SELECT perlet FROM grades_3 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY perlet "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("perlet")%>" <%if session("per90")=rs2("perlet") then response.write "selected"%>><%=rs2("perlet")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<!-- Periodo lançamentos -->
&nbsp;&nbsp;Lançamentos: <select size="1" name="perlanc" onchange="perlanc1()">
<option value="Todos" <%if session("lanc90")="Todos" then response.write "selected"%>>Todos</option>
<%
sql2="SELECT perlet2, right(perlet2,1) + 'º/' + left(perlet2,4) + ' ' + substring(perlet2,5,1) as descricao FROM grades_3 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY perlet2, right(perlet2,1) + 'º/' + left(perlet2,4) + substring(perlet2,5,1) " & _
"union all " & _
"SELECT left(perlet2,4) + '%' + right(perlet2,1), right(perlet2,1) + 'º/' + left(perlet2,4) FROM grades_3 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY left(perlet2,4) + '%' + right(perlet2,1), right(perlet2,1) + 'º/' + left(perlet2,4)"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("perlet2")%>" <%if session("lanc90")=rs2("perlet2") then response.write "selected"%>><%=rs2("descricao")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount
rs2.close
%>
</select>

<br>
Localizar por nome/chapa: <input type="text" name="localizar" size=35 value="<%=session("loc90")%>">
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