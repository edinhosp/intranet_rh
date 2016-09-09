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
sql="select id_plano, complementar, ordem, referencia from grades_plano_bi where id_plano>=0 order by id_plano "
rs.Open sql, ,adOpenStatic, adLockReadOnly
contador=1
do while not rs.eof 
'response.write "<br>" & len(rs("referencia")) & " : " & rs("referencia") & " : " & rs("id_plano")
if rs("complementar")=true then complementar=1 else complementar=0
if len(rs("referencia"))=0 then
	sql="delete from grades_plano_bi where id_plano=" & rs("id_plano") & " and ordem=" & rs("ordem") & " and complementar=" & complementar
	response.write "<font color=blue><br>" & contador & " " & sql & "<font color=black>"
	conexao.execute sql
	contador=contador+1
end if

referencia=""
ultimo=rs("id_plano")
rs.movenext
loop
rs.close

termino=now()
duracao=termino-inicio
response.write "<br>" & formatdatetime(now()-inicio,3)
response.write "<br>" & ultimo
%>

</body>
</html>