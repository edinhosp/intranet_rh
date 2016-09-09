<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
if session("UsuarioMaster")<>"02379" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Usuarios</title>
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<%
set conexao=server.createobject ("ADODB.Connection")
stringbd = "Provider=SQLOLEDB.1; SERVER=serveradm; DATABASE=corporerm; UID=sysdba; PWD=masterkey;"
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
%>
<img border="0" src="images/logo_centro_universitario_unifieo_big.jpg" width=225>
<table border="0" cellspacing="0" cellpadding="4">
<tr>
<td class=grupo valign="top">
<%
if session("usuariomaster")="02379" and Request.ServerVariables("REMOTE_ADDR")="10.0.1.91" then

sql="SELECT * From usuarios ORDER BY menu, nome"
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst
response.write "<table border='1' cellpadding='0' cellspacing='1' style='border-collapse: collapse' >"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulo><font size='1'>&nbsp;" & rs.fields(a).name & "</td>"
next
response.write "</tr>"
do while not rs.eof 
	response.write "<tr>"
	for a= 0 to rs.fields.count-1
		response.write "<td class=campo><font size='1'>&nbsp;" &rs.fields(a) & "</td>"
	next
	response.write "</tr>"
rs.movenext
loop
response.write "</table>"
rs.close
end if 'edson
%>
</td><td class=titulo valign="top">
<%
sql="SELECT l.*, nome From login l, usuarios u where u.usuario=l.usuario and isnull(saida) order by entrada"
sql="SELECT top 50 l.*, nome From login l left join usuarios u on u.usuario=l.usuario where saida is null order by entrada desc"
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	rs.movefirst
	response.write "<table border='1' cellpadding='0' cellspacing='1' style='border-collapse: collapse' >"
	response.write "<tr>"
	for a= 0 to rs.fields.count-1
		response.write "<td class=titulo><font size='1'>&nbsp;" & rs.fields(a).name & "</td>"
	next
	response.write "</tr>"
	do while not rs.eof 
		response.write "<tr>"
		for a= 0 to rs.fields.count-1
			response.write "<td class=campo><font size='1'>&nbsp;" &rs.fields(a) & "</td>"
		next
		response.write "</tr>"
	rs.movenext
	loop
	response.write "</table>"
end if
rs.close
%>
<p class=titulo>&nbsp;<%=application("usuariosativos")%>
</td></tr></table>
<%
for each strItem in Request.Cookies
	Response.write stritem & " = " & request.cookies(stritem) & "<br>"
	if request.cookies(stritem).haskeys then
		for each strsubitem in request.cookies(stritem)
		response.write "->" & stritem & "(" & strsubitem & ") = " & _
		request.cookies(stritem)(strsubitem) & "<br>"
		next
	end if
next
%>

<%
for each nome in Request.ServerVariables
	Response.write nome & " = " & Request.ServerVariables(nome) & "<br>"
next
%>
<br>
<%
for each nome in Session.Contents
	Response.write nome & " = " & Session.Contents(nome) & "<br>"
next
%>

<%
'sql="SELECT * from iamempresa"
'rs.Open sql, ,adOpenStatic, adLockReadOnly
'for a=0 to rs.fields.count-1
'response.write rs.fields(a).name & ": " & rs.fields(a) & " (" & rs.fields(a).type  & ")" & "<br>"
'next
'rs.close
%>

</body>

</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>