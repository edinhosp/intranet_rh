<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a60")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Geração de Arquivo de Teto</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
tetoinss=application("tetoinss")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form<>"" then
	ano=request.form("ano")
	mes=request.form("mes")
	valor=request.form("teto")
	valor=nraccess(valor)
	mesq=numzero(mes,2)
	'udia=day(dateserial(ano,mes+1,1)-1)
	conexao.execute "delete from tttetoinss"
	if mes="13" then
		evbase="297"
		evinss="071"
		tipo  ="04"
	else
		evbase="294"
		evinss="071"
		tipo  ="01"
	end if
	
	sql="INSERT INTO tttetoinss (CHAPA,CODEVE,ref) SELECT chapa, '" & evbase & "' AS Expr1, " & valor & " AS valor " & _
	"FROM rhcontroleteto WHERE ano='" & ano & "' AND mes=" & mes & " AND proporcional=0 "
	conexao.execute sql

	sql="INSERT INTO tttetoinss (CHAPA,CODEVE,ref) SELECT chapa, '" & evbase & "' AS Expr1, " & valor & " AS valor " & _
	"FROM rhcontroleteto WHERE ano='" & ano & "' AND mes=" & mes & " AND proporcional<>0 "
	conexao.execute sql

	sql="INSERT INTO tttetoinss (CHAPA,CODEVE,ref) SELECT chapa, '" & evinss & "' AS Expr1, proporcional " & _
	"FROM rhcontroleteto WHERE ano='" & ano & "' AND mes=" & mes & " AND proporcional<>0 "
	conexao.execute sql
end if
%>

<p class=titulo>Geração de arquivo do Controle de Teto do INSS para o RM Labore
<%
if request.form="" then 
%>
<form method="POST" action="inss_labore.asp" name="form">
<table border="1" bordercolor="#CCCCCC" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=200>
<tr>
	<td class=grupo>Ano</td>
	<td class=grupo>Mês</td>
	<td class=grupo>Teto INSS</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="ano" size="6" value="<%=year(now())%>"></td>
	<td class=titulo><input type="text" name="mes" size="4" value="<%=month(now())%>"></td>
	<td class=titulo><input type="text" name="teto" size="4" value="<%=tetoinss%>"></td>
</tr>
<tr>
	<td class=titulo colspan=3><input type="submit" value="Gerar arquivo" name="Gerar" class="button">
	</td>
</tr>
</table>
</form>
<%
else
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="tetoinss" & request.form("ano") & mesq & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	inicio=now()
	sql="SELECT CHAPA, CODEVE, ref FROM tttetoinss ORDER BY chapa, CODEVE "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	rs.movefirst
	do while not rs.eof 
		chapa=espaco1(rs("chapa"),16)
		evento=espaco1(rs("codeve"),4)
		valor=espaco1(replace(formatnumber(rs("ref"),2),".",""),15)
		leitura.writeline chapa & ";" & evento & ";" & valor & ";001;" & tipo
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
<a href="..\temp\<%=nomefile%>"><img src="../images/Diskette.gif" width="16" height="16" border="0" alt="">Arquivo Teto INSS</a>

<%
sql="SELECT t.CHAPA, f.NOME, t.CODEVE, t.ref " & _
"FROM tttetoinss t LEFT JOIN corporerm.dbo.PFUNC f ON t.CHAPA = f.CHAPA collate database_default " & _
"ORDER BY f.NOME, t.CODEVE "
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

end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%> 
</body>
</html>