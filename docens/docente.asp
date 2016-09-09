<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a5")="N" or session("a5")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.open Application("conexao")
'set conexao2=server.createobject ("ADODB.Connection")
'conexao2.open Application("consql")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form("B2")<>"" then
 	Session("PrimeiraVez")="Sim"
	if request.form("secao")="" then session("sel05")="Todas" else session("sel05")=request.form("secao")
	if request.form("situacao")="" then session("req05")="Todos" else session("req05")=request.form("situacao")

	if request.form("localizar")="" then session("loc05")="" else session("loc05")=request.form("localizar")

	if isnumeric(session("loc05"))=true then session("loc05")=numzero(session("loc05"),5)

	if session("sel05")<>"Todas" then
		session("sql05b")="AND (f.codsecao='" & session("sel05") & "') "
	else
		session("sql05b")=""
	end if

	if session("req05")="Todos" then
		session("sql05e")=""
	elseif session("req05")="Ativos" then
		session("sql05e")="AND (f.codsituacao in ('A','F','Z')) "
	elseif session("req05")="Afastados" then
		session("sql05e")="AND (f.codsituacao in ('E','I','L','M','O','P','R','T','U')) "
	elseif session("req05")="Demitidos" then
		session("sql05e")="AND (f.codsituacao in ('D','V','X')) "
	elseif session("req05")="RT" then
		sql1="SELECT CHAPA FROM corporerm.dbo.pfsalcmp WHERE CODEVENTO In ('255','256','257','258','128','138') GROUP BY CHAPA"
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		selecao="("
		do while not rs.eof
			selecao=selecao & "'" & rs("chapa") & "'"
			if rs.absoluteposition<rs.recordcount then selecao=selecao & ","
		rs.movenext:loop
		selecao=selecao & ")"
		session("sql05e")="AND f.chapa in " & selecao & " "
	elseif session("req05")="EFA" then
		sql1="SELECT CHAPA FROM grades_rt WHERE CODEVENTO In ('246','247') GROUP BY CHAPA"
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		selecao="("
		do while not rs.eof
			selecao=selecao & "'" & rs("chapa") & "'"
			if rs.absoluteposition<rs.recordcount then selecao=selecao & ","
		rs.movenext:loop
		selecao=selecao & ")"
		session("sql05e")="AND f.chapa in " & selecao & " "
	elseif session("req05")="RHT" then
		sql1="SELECT CHAPA FROM corporerm.dbo.pfsalcmp WHERE CODEVENTO In ('RHT') GROUP BY CHAPA"
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		selecao="("
		do while not rs.eof
			selecao=selecao & "'" & rs("chapa") & "'"
			if rs.absoluteposition<rs.recordcount then selecao=selecao & ","
		rs.movenext:loop
		selecao=selecao & ")"
		session("sql05e")="AND f.chapa in " & selecao & " "
	elseif session("req05")="Exceções" then
		sql1="SELECT CHAPA FROM grades_rt WHERE CODEVENTO not In ('255','256','257','258','128','246','247','275') GROUP BY CHAPA"
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		selecao="("
		do while not rs.eof
			selecao=selecao & "'" & rs("chapa") & "'"
			if rs.absoluteposition<rs.recordcount then selecao=selecao & ","
		rs.movenext:loop
		selecao=selecao & ")"
		session("sql05e")="AND f.chapa in " & selecao & " "
	end if
		
	if session("loc05")<>"" then
		if isnumeric(session("loc05")) then
			session("sql05d")="AND (f.chapa like '%" & session("loc05") & "%') "
	   else
			session("sql05d")="AND (f.nome like '%" & session("loc05") & "%') "
		end if
	else
		session("sql05d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

sqla="select f.chapa, f.nome, situacao=case when month(datademissao)=month(getdate()) and year(datademissao)=year(getdate()) then 'Ativo' else s.descricao end, " & _
"i.descricao as titulacao , c.titulacaopaga, i2.descricao as titulacao2 " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.pcodsituacao s, corporerm.dbo.pcodinstrucao i, corporerm.dbo.pfcompl c, corporerm.dbo.pcodinstrucao i2 " & _
"where (f.codpessoa=p.codigo and f.codsituacao=s.codcliente and p.grauinstrucao=i.codcliente and f.chapa=c.chapa and c.titulacaopaga=i2.codcliente " & _
"and /*f.chapa<'10000' and*/ f.codsindicato='03' " & session("sql05b") & session("sql05d") & session("sql05e") & ") " & _
"or (f.codpessoa=p.codigo and f.codsituacao=s.codcliente and p.grauinstrucao=i.codcliente and f.chapa=c.chapa and c.titulacaopaga=i2.codcliente " & _
"and f.chapa in ('00374','00257','00542','01129','01513','01514','00061','00057','00056','02297','02127') and f.codsindicato='01' " & session("sql05b") & session("sql05d") & session("sql05e") & ") "
'"and f.chapa='00374' and f.codsindicato='01' " & session("sql05b") & session("sql05d") & session("sql05e") & ") "
sqlc="order by f.nome"
sql1=sqla & session("sql05b") & session("sql05d") & session("sql05e") & sqlc
sql1=sqla & sqlc

if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	rs2.CacheSize = registros
	rs2.PageSize = registros
	'set rs.ActiveConnection = conexao
	rs2.Open sql1, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez")="Nao"
else
	if request("folha")="" then pagina=1
	if request.form("pagina")<>"" then pagina=request.form("pagina")
	if request("folha")<>"" then pagina=request("folha")
	Session("Pagina")=pagina
	conexao.cursorlocation = 3 'aduseclient
	rs2.CacheSize = registros
	rs2.PageSize = registros
	'set rs2.ActiveConnection = conexao
	rs2.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then 	MostraDados
end if	
	
Sub MostraDados()
	if rs2.recordcount>0 then 	rs2.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
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
for selpag=1 to rs2.pagecount
	if selpag=atual then selpag1="selected" else selpag1=""
	response.write "<option value=" & selpag & " " & selpag1 & ">" & selpag & "</option>"
next
response.write "</select>"
response.write "/" & rs2.pagecount & "</b>&nbsp;"

if atual=rs2.pagecount or rs2.pagecount=0 then
	response.write "<img src='../images/setanext0.gif' border='0'>"
	response.write "<img src='../images/setalast0.gif' border='0'>"
else
	response.write "<a href=""docente.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
	response.write "<a href=""docente.asp?folha=" & rs2.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="30%" valign="top" align="right">
<%
Response.write "Registros: " & rs2.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="690" cellpadding="2" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Situação</td>
	<td class=titulo align="center">Tit.Paga/MEC</td>
	<td class=titulo align="center">Tit.Real</td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs2.recordcount>0 then
For Contador=1 to Session("RegistrosPorPagina")
nome=rs2("nome"):'if rs2("chapa")="00362" then nome="-"
%>
<tr>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    <a class=r href="docente_ver.asp?chapa=<%=rs2("chapa")%>&nome=<%=nome%>" onclick="NewWindow(this.href,'CadastroProfessor','645','480','yes','center');return false" onfocus="this.blur()">
	&nbsp;<%=rs2("chapa")%></a></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    &nbsp;<%=rs2("nome")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    <%=rs2("situacao")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    &nbsp;<%=rs2("titulacao2")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    &nbsp;<%=rs2("titulacao")%></td>
</tr>
<%
	if linha=1 then linha=0 else linha=1
	rs2.movenext
	if rs2.eof then exit for
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
'sql2="SELECT f.CODSECAO, S.DESCRICAO FROM dc_professor as f, PSECAO as s " & _
'"WHERE f.CODSECAO = S.CODIGO GROUP BY f.CODSECAO, s.DESCRICAO order by s.descricao"
sql2="select f.codsecao, s.descricao from corporerm.dbo.pfunc f, corporerm.dbo.psecao s " & _
"where f.codsecao=s.codigo and f.codsindicato='03' and f.chapa<'10000' " & _
"group by f.codsecao, s.descricao order by s.descricao "
if session("grupoacesso")="10" or session("grupoacesso")="55" or session("usuariogrupo")="CHEFE DEPTO" then
	'sql2="SELECT f.CODSECAO, S.DESCRICAO FROM dc_professor as f, PSECAO as s, " & _
	'"(select chapa1 from g2ch where codcur in (select codcur from grades_user where usuario='" & session("usuariomaster") & "') group by chapa1) as g  " & _
	'"WHERE f.CODSECAO = S.CODIGO AND f.chapa=g.chapa1 " & _
	'"GROUP BY f.CODSECAO, s.DESCRICAO order by s.descricao"
	sql1="select chapa1 from g2ch where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') group by chapa1 "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	selecao="("
	do while not rs.eof
		selecao=selecao & "'" & rs("chapa1") & "'"
		if rs.absoluteposition<rs.recordcount then selecao=selecao & ","
	rs.movenext
	loop
	selecao=selecao & ")"
	sql2="select f.codsecao, s.descricao from corporerm.dbo.pfunc f, corporerm.dbo.psecao s " & _
	"where f.codsecao=s.codigo and f.codsindicato='03' and f.chapa in " & selecao & " " & _
	"group by f.codsecao, s.descricao order by s.descricao "
end if
rs3.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<p><font size=1>Filtrar Seção: <select size="1" name="secao">
<option value="Todas">Todas Seções</option>
<%
rs3.movefirst: do while not rs3.eof
%>
    <option value="<%=rs3("codsecao")%>" <%if session("sel05")=rs3("codsecao") then response.write "selected"%>><%=rs3("codsecao") & " - " & rs3("descricao")%></option>
<%
rs3.movenext: loop
%>
</select>

Filtrar Situação: <select size="1" name="situacao">
<option value="Todos" <%if session("req05")="Todos" then response.write "selected"%>>Todas Situações</option>
    <option value="Ativos" <%if session("req05")="Ativos" then response.write "selected"%>>Ativos</option>
    <option value="Afastados" <%if session("req05")="Afastados" then response.write "selected"%>>Afastados</option>
    <option value="Demitidos" <%if session("req05")="Demitidos" then response.write "selected"%>>Demitidos</option>
    <option value="RT" <%if session("req05")="RT" then response.write "selected"%>>RT</option>
    <option value="EFA" <%if session("req05")="EFA" then response.write "selected"%>>EFA</option>
    <option value="RHT" <%if session("req05")="RHT" then response.write "selected"%>>RHT</option>
    <option value="Exceções" <%if session("req05")="Exceções" then response.write "selected"%>>Exceções</option>
</select>

<br>
Localizar por nome/chapa: <input type="text" name="localizar" size=25 value="<%=session("loc05")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar"></p>
</form>

<%
rs2.close
set rs=nothing
rs3.close
set rs2=nothing
conexao.close
set conexao=nothing
'conexao2.close
'set conexao2=nothing
%>
</body>
</html>