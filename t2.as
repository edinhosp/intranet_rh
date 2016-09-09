<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=false
	Server.ScriptTimeout = 16000
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Plano de Ensino - Relatório</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, rs3
set conexao=server.createobject ("ADODB.Connection")
conexao.open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
%>

<%
inicio=now()
sql="select id_plano, complementar, ordem, referencia from grades_plano_bi where referencia is not null and id_plano>=0 order by id_plano "
rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof 
response.write "<br>" & rs("referencia") & ": "
for a=1 to len(rs("referencia"))
	letra=mid(rs("referencia"),a,1)
	codigo=asc(letra)
	if codigo=13 or codigo=10 then
		referencia=referencia
	else
		referencia=referencia+letra
	end if
	if a=len(rs("referencia")) then
	end if
next
		'if a=len(rs("bibliografia")) then referencia=referencia+letra
if rs("complementar")=true then complementar=1 else complementar=0
sql2="update grades_plano_bi set referencia='" & referencia & "' where id_plano=" & rs("id_plano") & " and ordem=" & rs("ordem") & " and complementar=" & complementar & ""
if rs.absoluteposition=1 then response.write sql2
response.write " " & rs.absoluteposition & "/" & rs.recordcount & "-" & rs("id_plano")
conexao.execute sql2
referencia=""

referencia=""
ultimo=rs("id_plano")
rs.movenext
loop
rs.close

termino=now()
duracao=termino-inicio
response.write "<br><font size=4>" & formatdatetime(now()-inicio,3)
response.write "<br>" & ultimo
%>

</body>
</html>