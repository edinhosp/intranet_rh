<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a64")="N" or session("a64")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Bolsas de Estudo</title>
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
registros=Session("RegistrosPorPagina")
registros=250
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")
conexao.open Application("conexao")
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("secao")="" then session("sel64")="Todas" else session("sel64")=request.form("secao")
	if request.form("tipocurso")="" then session("tcur64")="Todos" else session("tcur64")=request.form("tipocurso")
	if request.form("tipobolsa")="" then session("tbolsa64")="Todos" else session("tbolsa64")=request.form("tipobolsa")
	if request.form("situacao")="" then session("sit64")="Todos" else session("sit64")=request.form("situacao")
	if request.form("curso")="" then session("cur64")="Todos" else session("cur64")=request.form("curso")
	if request.form("situacaof")="" then session("sitf64")="Todos" else session("sitf64")=request.form("situacaof")
	if request.form("localizar")="" then session("loc64")="" else session("loc64")=request.form("localizar")
	if isnumeric(session("loc64"))=true then session("loc64")=numzero(session("loc64"),5)

	if session("sitf64")<>"Todos" then
		session("sql64h")="and (f.codsituacao='" & session("sitf64") & "') "
	else
		session("sql64h")=""
	end if

	if session("cur64")<>"Todos" then
		session("sql64g")="and (b.curso='" & session("cur64") & "') "
	else
		session("sql64g")=""
	end if

	if session("sit64")<>"Todos" then
		session("sql64f")="and (b.situacao='" & session("sit64") & "') "
	else
		session("sql64f")=""
	end if

	if session("tcur64")<>"Todos" then
		session("sql64c")="and (b.tipocurso='" & session("tcur64") & "') "
	else
		session("sql64c")=""
	end if

	if session("tbolsa64")<>"Todos" then
		session("sql64e")="and (b.tp_bolsa='" & session("tbolsa64") & "') "
	else
		session("sql64e")=""
	end if

	if session("sel64")<>"Todas" then
		session("sql64b")="AND (f.codsecao='" & session("sel64") & "') "
	else
		session("sql64b")=""
	end if

	if session("loc64")<>"" then
   		if isnumeric(session("loc64")) then
			session("sql64d")="AND (b.chapa like '%" & session("loc64") & "%') "
		else
			session("sql64d")="AND (f.nome like '%" & session("loc64") & "%' or b.nome_bolsista like '%" & session("loc64") & "%') "
		end if
	else
		session("sql64d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

sqla="select b.chapa, f.nome, count(b.chapa) as total, f.codsecao, s.descricao " & _
"from corporerm.dbo.pfunc f, bolsistas b, corporerm.dbo.psecao s " & _
"where f.chapa collate database_default=b.chapa and f.codsecao=s.codigo "
sqlb=" "
sqlc="order by f.nome, b.nome_bolsista "
sqla="select b.id_bolsa, b.chapa, f.nome, b.nome_bolsista, b.parentesco, b.curso, s.descricao as situacao, t.descricao as tipo " & _
", f.codsecao, se.descricao as secao, b.tipocurso, s.id_sit, b.tp_bolsa, b.situacao as codsit, f.codsituacao " & _
"from corporerm.dbo.pfunc f, bolsistas b, bolsistas_situacao s, bolsistas_tipo t, corporerm.dbo.psecao se " & _
"where f.chapa collate database_default=b.chapa and b.situacao=s.id_sit and b.tp_bolsa=t.id_tp and f.codsecao=se.codigo "

sql1=sqla & sqlb & session("sql64b") & session("sql64c") & session("sql64d") & session("sql64e") & session("sql64f") & session("sql64g") & session("sql64h") & sqlc

if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	Set rs.ActiveConnection = conexao
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
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	Set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 	MostraDados
end if	

Sub MostraDados()
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
End Sub
%>

<form method="POST" name="form" action="bolsa.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Beneficiários de Bolsas de Estudos</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="55%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""bolsa.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""bolsa.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
end if

response.write "&nbsp;<b>"
response.write "<select size='1' name='pagina' onchange='javascript:submit()'>"
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
response.write "<a href=""bolsa.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""bolsa.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="30%" valign="top" align="center">
<% if session("a64")="T" then %>
<a href="bolsa_nova.asp" onclick="NewWindow(this.href,'Inclusao','520','320','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif">
<font size="1">inserir novo beneficiário</font></a>
<% end if %>

	</td>
	<td class=campo width="15%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="690" cellpadding="1" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Titular</td>
	<td class=titulo align="center">Nome Bolsista</td>
	<td class=titulo align="center">Tipo<br>Curso</td>
	<td class=titulo align="center">Tipo Bolsa<br>Situação</td>
	<td class=titulo align="center">&nbsp;</td>
	<td class=titulo align="center">&nbsp;</td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
	<td class="campor">
	<a href="titular_ver.asp?codigo=<%=rs("chapa")%>" onclick="NewWindow(this.href,'Alteracao','690','500','yes','center');return false" onfocus="this.blur()">
		<font size="1"><%=rs("chapa")%></font></a>-<%=rs("nome")%>
	</td>
	<td class="campor"><%=rs("nome_bolsista") %></td>
	<td class="campor"><font color=blue><%=rs("tipocurso")%></font><br><%=rs("curso") %></td>
	<td class="campor"><b><font color="#660099"><%=rs("tipo")%></b><br><font color=#0000CC><%if rs("id_sit")="M" then response.write "<b>"%><%=rs("situacao")%></td>
	<td class="campor" align="center">
	<% if session("a64")="T" or session("a64")="C" then %>
		<a href="bolsa_alteracao.asp?codigo=<%=rs("id_bolsa")%>" onclick="NewWindow(this.href,'AlteracaoBolsista','510','330','no','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0" width=13 alt="Alterar os dados do bolsista"></a>
	<% end if %>
	</td>
	<td class="campor" align="center">
	<% if session("a64")="T" or session("a64")="C" then %>
		<a href="bolsa_ver.asp?codigo=<%=rs("id_bolsa")%>" onclick="NewWindow(this.href,'BolsaVer','690','600','no','center');return false" onfocus="this.blur()">
		<img src="../images/Form.gif" width="16" height="16" border="0" alt="Alterar os dados da bolsa de estudo"></a>
	<% end if %>
	</td>
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
<p>Esta seleção não mostra nenhum registro.</p>
<%
end if
%>
<hr>
<p style="margin-bottom:0;margin-top:0"><font size=1>Filtrar Seção: <select size="1" name="secao" class=a>
<option value="Todas">Todas Seções</option>
<%
sql2="SELECT S.CODIGO, S.DESCRICAO FROM BOLSISTAS B, corporerm.dbo.PFUNC F, corporerm.dbo.PSECAO S " & _
"WHERE B.CHAPA=F.CHAPA collate database_default AND F.CODSECAO = S.CODIGO " & _
"GROUP BY S.CODIGO, S.DESCRICAO order by S.descricao"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
    <option value="<%=rs2("codigo")%>" <%if session("sel64")=rs2("codigo") then response.write "selected"%>><%=rs2("codigo") & " - " & rs2("descricao")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
<font size=1>Tipo Curso: <select size=1 name="tipocurso" class=a>
<option value="Todos">Todos tipos</option>
<%
sql2="SELECT tipocurso from bolsistas group by tipocurso"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
    <option value="<%=rs2("tipocurso")%>" <%if session("tcur64")=rs2("tipocurso") then response.write "selected"%>><%=rs2("tipocurso")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
<p style="margin-bottom:0;margin-top:5"><font size=1>Tipo Bolsa: <select size=1 name="tipobolsa" class=a>
<option value="Todos">Todos tipos</option>
<%
sql2="SELECT b.tp_bolsa, t.descricao FROM bolsistas b, bolsistas_tipo t WHERE b.tp_bolsa=t.id_tp " & _
"GROUP BY b.tp_bolsa, t.descricao ORDER BY t.descricao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
    <option value="<%=rs2("tp_bolsa")%>" <%if session("tbolsa64")=rs2("tp_bolsa") then response.write "selected"%>><%=rs2("descricao")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>

<font size=1>Situação: <select size=1 name="situacao" class=a>
<option value="Todos">Todos situações</option>
<%
sql2="SELECT b.situacao, t.descricao FROM bolsistas b, bolsistas_situacao t WHERE b.situacao=t.id_sit " & _
"GROUP BY b.situacao, t.descricao ORDER BY t.descricao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
    <option value="<%=rs2("situacao")%>" <%if session("sit64")=rs2("situacao") then response.write "selected"%>><%=rs2("descricao")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>

<font size=1>Curso: <select size=1 name="curso" class=a>
<option value="Todos">Todos cursos</option>
<%
sql2="SELECT curso from bolsistas group by curso "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
    <option value="<%=rs2("curso")%>" <%if session("cur64")=rs2("curso") then response.write "selected"%>><%=rs2("curso")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>

<p style="margin-bottom:0;margin-top:5"><font size=1>
Situação: <select size=1 name="situacaof" class=a>
<option value="Todos">Todos Func.</option>
<%
sql2="SELECT f.codsituacao, s.descricao FROM bolsistas b, corporerm.dbo.pfunc f, corporerm.dbo.pcodsituacao s WHERE b.chapa=f.chapa collate database_default and f.codsituacao=s.codcliente " & _
"GROUP BY f.codsituacao, s.descricao ORDER BY f.codsituacao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
    <option value="<%=rs2("codsituacao")%>" <%if session("sitf64")=rs2("codsituacao") then response.write "selected"%>><%=rs2("descricao")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>

Localizar por nome/chapa: <input type="text" name="localizar" size=35 value="<%=session("loc64")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
<input name="B2" type="submit" class="button" value="Clique para Filtrar">
</font></form>

</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>