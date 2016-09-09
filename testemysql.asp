<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.redirect "intranet.asp"
if session("a1")="N" or session("a1")="" then response.redirect "intranet.asp"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Teste MySQL</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->

<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("mysqlfieo")
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=localhost; Port=3306; Option=0; Socket=; Stmt=; Database=rhonline2; Uid=root; Pwd="
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=colossus2.fieo.br; Port=3306; Option=0; Socket=; Stmt=; Database=website; Uid=rh; Pwd=!@#qaz"

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

sqla="SELECT * from uni_rh_curriculum limit 10"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
'*************** inicio teste **********************
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
response.write "<p>"
'*************** fim teste **********************%>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>

<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>