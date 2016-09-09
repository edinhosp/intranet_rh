<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a37")="N" or session("a37")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastro UNIFIEO</title>
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
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rsc=server.createobject ("ADODB.Recordset")
set rsc.ActiveConnection = conexao

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("localizar")="" then
		session("loc37")=""
	else
		session("loc37")=request.form("localizar")
	end if
	'if session("loc424")<>"" then
		if isnumeric(session("loc37")) then
			session("sql37d")="where (codigo like '%" & session("loc37") & "%') "
		else
			session("sql37d")="where (nome like '%" & session("loc37") & "%') "
		end if
	'else
		'session("sqld424")=""
	'end if

	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if

sqla="SELECT CODIGO, NOME, TIPO, STATUS FROM (select chapa codigo, nome, tipo=case when codtipo='T' then 'ESTAGIÁRIO' when codsindicato='03' then 'PROFESSOR' else 'ADMINISTRATIVO' end, s.descricao as status from corporerm.dbo.pfunc f, corporerm.dbo.pcodsituacao s where (f.chapa<'10000' or f.chapa>'90000') and f.codsituacao=s.codcliente " & _
"UNION ALL " & _
"SELECT E.MATRICULA, E.NOME, 'ALUNO', S.DESCRICAO FROM corporerm.dbo.EALUNOS E LEFT JOIN corporerm.dbo.USITMAT S ON E.STATUS=S.CODSITMAT) AS QUERY " 
sqlb=""
sqlc="order by nome "
sqla="SELECT CODIGO, NOME, TIPO, STATUS FROM (select chapa codigo, nome, tipo=case when codtipo='T' then 'ESTAGIÁRIO' when codsindicato='03' then 'PROFESSOR' else 'ADMINISTRATIVO' end, s.descricao as status from corporerm.dbo.pfunc f, corporerm.dbo.pcodsituacao s where (f.chapa<'10000' or f.chapa>'90000') and f.codsituacao=s.codcliente ) AS QUERY "

sql1=sqla & sqlb & session("sql37d") & sqlc

if session("a37")="E" then sql1="SELECT E.MATRICULA as codigo, E.NOME, 'ALUNO' as tipo, S.DESCRICAO as status FROM corporerm.dbo.EALUNOS E LEFT JOIN corporerm.dbo.USITMAT S ON E.STATUS=S.CODSITMAT " & session("sql37d") & " order by nome "

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
<form method="POST" action="pessoas2.asp" name="form">
<input type="hidden" name="vez1" value="<%=session("PrimeiraVez")%>">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro - UNIFIEO</p>
<p style="margin-top: 0; margin-bottom: 0"><font color="blue">
Localizar por nome/chapa: <input type="text" name="localizar" size=35 value="<%=session("loc37")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar">

<table border="0" width="650" cellspacing="0" style="border-collapse: collapse" cellpadding="0">
<tr>
    <td class=campo width="60%" valign="center" align="left">Página: 
<%
Session("Load1")="1"
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""pessoas2.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""pessoas2.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
response.write "<a href=""pessoas2.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""pessoas2.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
    <td class=campo width="20%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
	<td class=campo width="20%" valign="top" align="right">
	</td>
  </tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="650" cellpadding="1" style="border-collapse: collapse">
<tr>
    <td class=titulo align="center">Chapa/Matrícula</td>
    <td class=titulo align="center">Nome</td>
    <td class=titulo align="center">Tipo</td>
    <td class=titulo align="center">Status</td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
    <td class=campo><%=rs("codigo")%></td>
	<td class=campo>
<% if rs("tipo")="ALUNO" then %>	
    <a href="alunos_ver.asp?matricula=<%=rs("codigo")%>" onclick="NewWindow(this.href,'Ouvidoria','690','500','yes','center');return false" onfocus="this.blur()">
	<%=rs("nome")%></a>
<% else %>
    <a href="funcionarios_ver.asp?chapa=<%=rs("codigo")%>" onclick="NewWindow(this.href,'Ouvidoria','690','500','yes','center');return false" onfocus="this.blur()">
	<%=rs("nome")%></a>
<% end if %>
	</td>
    <td class=campo><%=rs("tipo")%></td>
    <td class=campo><%=ucase(rs("status"))%></td>
</tr>
<%
if linha=1 then linha=0 else linha=1
rs.movenext
if rs.eof then exit for
'loop
Next

else 'sem registros
%>
<td class=grupo colspan=10>Esta seleção não mostra nenhum registro.</td>
<%
end if

rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</table>

</form>
</body>
</html>