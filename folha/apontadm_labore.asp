<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Geração de Arquivo de Lançamentos</title>
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

if request.form<>"" then
	ano=request.form("ano")
	mes=request.form("mes")
	mesq=numzero(mes,2)
	udia=day(dateserial(ano,mes+1,1)-1)
	conexao.execute "delete from ttcodigofixo where sessao='" & session("usuariomaster") & "' "
	
	sql="INSERT INTO ttcodigofixo ( sessao, CHAPA, NOME, CODEVE, descricao, TotalCH, horas, ref, fator ) " & _
	"SELECT '" & session("usuariomaster") & "', a.CHAPA, f.NOME, a.CODEVENTO, e.DESCRICAO, " & _
	"tplanc=case when valhordiaref='H' then 1 else case when valhordiaref='D' then 2 else 0 end end, " & _
	"base=case when valhordiaref='H' then valor/60/24 else valor end, " & _
	"a.VALOR, a.NROVEZES " & _
	"FROM (apont_adm a INNER JOIN corporerm.dbo.PFUNC f ON a.CHAPA=f.CHAPA collate database_default) INNER JOIN apont_adm_eventos e ON a.CODEVENTO = e.CODIGO " & _
	"WHERE a.ano='" & ano & "' AND a.mes=" & mes
	conexao.execute sql

	sql="SELECT t.CHAPA, t.NOME, t.CODEVE, t.descricao, Sum(t.horas) AS horas, Sum(t.ref) AS ref " & _
	"FROM ttcodigofixo t WHERE t.sessao='" & session("usuariomaster") & "' " & _
	"GROUP BY t.CHAPA, t.NOME, t.CODEVE, t.descricao " & _
	"ORDER BY t.CODEVE, t.CHAPA, t.CODEVE;"
	sql="SELECT t.CHAPA, t.NOME, t.CODEVE, t.descricao, t.ref as horas, t.ref AS ref " & _
	"FROM ttcodigofixo t WHERE t.sessao='" & session("usuariomaster") & "' " & _
	"ORDER BY t.CODEVE, t.CHAPA "
end if

if request.form="" then
%>
<p class=titulo>Geração de arquivo de Lançamentos para Folha
<form method="POST" action="apontadm_labore.asp">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Ano</td>
	<td class=titulo>Mês</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="ano" size="6" value="<%=year(now())%>"></td>
	<td class=titulo><input type="text" name="mes" size="4" value="<%=month(now())%>" class=a></td>
</tr>
<tr>
	<td class=titulo colspan="2"><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></td>
</tr>
</table>
</form>
<% else %>
<%
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="lancamentos" & request.form("ano") & mesq & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	inicio=now()
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		chapa=espaco1(rs("chapa"),16)
		evento=espaco1(rs("codeve"),4)
		valor=espaco1(replace(formatnumber(rs("ref"),2),".",""),15)
		leitura.writeline chapa & ";" & evento & ";" & valor & ";001;01"
	rs.movenext
	loop
	rs.close
	termino=now()
	duracao=(termino-inicio)
	'Response.write "Inicio: " & inicio & "<br>Termino: " & termino & "<br>Duracao: " & formatdatetime(duracao,3)
	leitura.close
	set leitura=nothing
	set arquivo=nothing

%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width='650'>
<tr>
	<td><p class=titulo>Geração de arquivo de Lançamentos para Folha</td>
	<td><a href="../temp/<%=nomefile%>">Arquivo Lançamentos</a></td>
</tr>
</table>
<%
rs.Open sql, ,adOpenStatic, adLockReadOnly
total=0
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse' width='650'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulo>" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	conteudo=rs.fields(a)
	if a=4 then
	if rs.fields(2)="088" or rs.fields(2)="089" or rs.fields(2)="003" or rs.fields(2)="004" or rs.fields(2)="010" then
		temp=cdbl(rs.fields(a))
		hora=int(temp/60)
		minuto=temp-(hora*60)
		conteudo=hora& ":" & numzero(minuto,2)
		'conteudo
	else
		conteudo=rs.fields(a)
	end if
	end if
	response.write "<td class=campo>" & conteudo & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"

end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%> 
</body>
</html>