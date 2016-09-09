<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Geração de Arquivo</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

ipacesso=Request.ServerVariables("REMOTE_ADDR")
if ipacesso="10.0.1.91" or ipacesso="10.0.1.10" then
%>

<p class=titulo>Geração de arquivo
<%
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="juliano" & day(now) & month(now) & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	inicio=now()
sql="SELECT f.CHAPA, f.NOME, p.EMAIL, u.CURSO " & _
"FROM (pfunc AS f INNER JOIN ppessoa AS p ON f.CODPESSOA = p.CODIGO) LEFT JOIN uprofformacao_ AS u ON f.CHAPA = u.CODPROF " & _
"WHERE f.CODSINDICATO='03' AND f.CODSITUACAO In ('A','F','Z') AND u.TIPO='Graduação'; "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	rs.movefirst
	do while not rs.eof 
		chapa=espaco2(rs("chapa"),5)
		nome=espaco2(rs("nome"),45)
		email=espaco2(rs("email"),60)
		curso=espaco2(rs("curso"),100)
		leitura.writeline chapa & ";" & nome & ";" & email & ";" & curso
	rs.movenext
	loop
	end if 'recordcount
	rs.close
	termino=now()
	duracao=(termino-inicio)
	'Response.write "Inicio: " & inicio & "<br>Termino: " & termino & "<br>Duracao: " & formatdatetime(duracao,3)
	leitura.close
	set leitura=nothing
	set arquivo=nothing

%>
<p style="margin-top: 0; margin-bottom: 0">
<a href="..\temp\<%=nomefile%>"><img src="../images/Diskette.gif" width="16" height="16" border="0" alt="">Arquivo</a>

<%
rs.Open sql, ,adOpenStatic, adLockReadOnly
total=0
if rs.recordcount>0 then
rs.movefirst
response.write "<table border='1' cellpadding='0' cellspacing='3' style='border-collapse: collapse' >"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td><font size='1'>&nbsp;" & rs.fields(a).name & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	if rs.fields(a).type=5 then conteudo=formatnumber(rs.fields(a),2) else conteudo=rs.fields(a)
	'response.write "<td><font size='1'>&nbsp;" &rs.fields(a) & rs.fields(a).type & "</td>"
	response.write "<td><font size='1'>&nbsp;" & conteudo & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
else
	response.write "<p><b>Não existem cartas de teto máximo cadastradas para este mês."
end if 'rs.recordcount
rs.close
response.write "<p>"

end if 'request ip

set rs=nothing
conexao.close
set conexao=nothing
%> 
</body>
</html>