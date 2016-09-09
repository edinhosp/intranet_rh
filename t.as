<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=false
	Server.ScriptTimeout = 160000
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
'18281
inicio=now()
numero=0
sql="select id_plano, biblio, complementar from (" & _
"select id_plano, bibliografiac biblio, complementar=1 from grades_plano where bibliografiac is not null " & _
"union all " & _
"select id_plano, bibliografia biblio, complementar=0 from grades_plano where bibliografia is not null ) z " & _
"where id_plano>" & numero & " order by id_plano "
'"where id_plano in (select id_plano from grades_plano_acerto) order by id_plano "

rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof 
ordem=1
for a=1 to len(rs("biblio"))
	letra=mid(rs("biblio"),a,1)
	codigo=asc(letra)
	if codigo<>10 or codigo<>13 then
		referencia=referencia+letra
	end if
	if codigo=10 or a=len(rs("biblio")) then
		'if a=len(rs("bibliografia")) then referencia=referencia+letra
		sql2="insert into grades_plano_bi (id_plano, complementar, ordem, referencia) " & _
		"select " & rs("id_plano") & ", " & rs("complementar") & ", " & ordem & ",'" & referencia & "'"
		'response.write "<br>" & sql2
		response.write "<br> " & rs.absoluteposition & "/" & rs.recordcount & " : " & rs("id_plano")
		conexao.execute sql2
		referencia=""
		ordem=ordem+1
	end if
next

referencia=""
rs.movenext
loop
rs.close

sql="select max(id_plano) ultimo from grades_plano_bi "
rs.Open sql, ,adOpenStatic, adLockReadOnly
response.write "<br><font size=3>" & rs("ultimo") & "<br>"
rs.close
termino=now()
duracao=termino-inicio
response.write formatdatetime(now()-inicio,3)
%>

</body>
</html>