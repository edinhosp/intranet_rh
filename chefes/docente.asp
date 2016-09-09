<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a98")="N" or session("a98")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Docentes</title>
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
<body topmargin="5" leftmargin="5">
<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")
conexao.open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")

if request.form("B2")<>"" then
 	Session("PrimeiraVez")="Sim"
	if request.form("secao")="" then session("sel98")="Todas" else session("sel98")=request.form("secao")
	if request.form("situacao")="" then session("req98")="Todos" else session("req98")=request.form("situacao")
	if request.form("localizar")="" then session("loc98")="" else session("loc98")=request.form("localizar")
	if isnumeric(session("loc98"))=true then session("loc98")=numzero(session("loc98"),5)

	if session("sel98")<>"Todas" then
		session("sql98b")=" and au.area='" & request.form("secao") & "' "
	else
		session("sql98b")=""
	end if

	if session("req98")="Todos" then
		session("sql98e")=""
	elseif session("req98")="Ativos" then
		session("sql98e")="AND (f.codsituacao in ('A','F','Z')) "
	elseif session("req98")="Afastados" then
		session("sql98e")="AND (f.codsituacao in ('E','I','L','M','O','P','R','T','U')) "
	elseif session("req98")="Demitidos" then
		session("sql98e")="AND (f.codsituacao in ('D','V','X')) "
	elseif session("req98")="RT" then
		session("sql98e")="AND f.chapa in (SELECT CHAPA FROM grades_rt WHERE CODEVENTO In ('255','256','257','258','128') and fim>now-30 GROUP BY CHAPA) "
	elseif session("req98")="EFA" then
		session("sql98e")="AND f.chapa in (SELECT CHAPA FROM grades_rt WHERE CODEVENTO In ('246','247') GROUP BY CHAPA) "
	elseif session("req98")="SAJ" then
		session("sql98e")="AND f.chapa in (SELECT CHAPA FROM grades_rt WHERE CODEVENTO In ('275') GROUP BY CHAPA) "
	elseif session("req98")="Exceções" then
		session("sql98e")="AND f.chapa in (SELECT CHAPA FROM grades_rt WHERE CODEVENTO not In ('255','256','257','258','128','246','247','275') GROUP BY CHAPA) "
	end if
		
	if session("loc98")<>"" then
		if isnumeric(session("loc98")) then
			session("sql98d")="AND (f.chapa like '%" & session("loc98") & "%') "
	   else
			session("sql98d")="AND (f.nome like '%" & session("loc98") & "%') "
		end if
	else
		session("sql98d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

sqla="select f.chapa, f.nome, iif(month(datademissao)=month(now) and year(datademissao)=year(now),'Ativo',s.descricao) as situacao, i.descricao as titulacao " & _
"from dc_professor_t f, pcodsituacao s, pcodinstrucao i, " & _
"(SELECT ap.chapa1 AS chapa FROM grades_areacon_p2 AS ap, grades_areacon_u AS au " & _
"where ap.area=au.area and au.usuario='" & session("usuariomaster") & "' " & session("sql98b") & " GROUP BY ap.chapa1) a " & _
"where f.codsituacao = s.codcliente and f.grauinstrucao=i.codcliente " & _
"and f.chapa=a.chapa "

sqlc="order by f.nome"
sql1=sqla & session("sql98d") & session("sql98e") & sqlc

if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
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
<form method="POST" name="form" action="docente.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">CADASTRO DE DOCENTES</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="70%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
	response.write "<img src='../images/setafirst0.gif' border='0'>"
	response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
	response.write "<a href=""docente.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
	response.write "<a href=""docente.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
	response.write "<a href=""docente.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
	response.write "<a href=""docente.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="30%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="650" cellpadding="2" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Situação</td>
	<td class=titulo align="center">Titulação</td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to Session("RegistrosPorPagina")
nome=rs("nome"):if rs("chapa")="00362" then nome="-"
%>
<tr>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    <a class=r href="docente_ver.asp?chapa=<%=rs("chapa")%>&nome=<%=nome%>" onclick="NewWindow(this.href,'CadastroProfessor','645','480','yes','center');return false" onfocus="this.blur()">
	&nbsp;<%=rs("chapa")%></a></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    &nbsp;<%=rs("nome")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    <%=rs("situacao")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    &nbsp;<%=rs("titulacao")%></td>
</tr>
<%
	if linha=1 then linha=0 else linha=1
	rs.movenext
	if rs.eof then exit for
'loop
Next
%>
</table>
<%
else 'sem registros
%>
<p><b><font color="#FF0000">Esta seleção não mostra nenhum registro.</font></b></p>
<%
end if 'sem registros
%>
<hr>
<%
sql2="select area from grades_areacon_u where usuario='" & session("usuariomaster") & "' order by area "
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<p><font size=1>Area de conhecimento: <select size="1" name="secao">
<option value="Todas">Todas áreas</option>
<%
rs2.movefirst: do while not rs2.eof
%>
    <option value="<%=rs2("area")%>" <%if session("sel98")=rs2("area") then response.write "selected"%>><%=rs2("area")%></option>
<%
rs2.movenext: loop
%>
</select>

Filtrar Situação: <select size="1" name="situacao">
<option value="Todos" <%if session("req98")="Todos" then response.write "selected"%>>Todas Situações</option>
    <option value="Ativos" <%if session("req98")="Ativos" then response.write "selected"%>>Ativos</option>
    <option value="Afastados" <%if session("req98")="Afastados" then response.write "selected"%>>Afastados</option>
    <option value="Demitidos" <%if session("req98")="Demitidos" then response.write "selected"%>>Demitidos</option>
</select>

<br>
Localizar por nome/chapa: <input type="text" name="localizar" size=25 value="<%=session("loc98")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar"></p>
</form>

<%
rs.close
set rs=nothing
rs2.close
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>