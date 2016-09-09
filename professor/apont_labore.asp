<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a57")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Geração de Arquivo do Apontamento dos Professores</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
inicio=now()
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
sessao=session.sessionid

if request.form<>"" then
	mesbase=request.form("mesbase")
'	conexao.execute "delete * from ttcodigofixo"

sql="DELETE FROM apontamento_arquivo WHERE mes_base='" & dtaccess(mesbase) & "' and sessao='" & session("usuariomaster") & "' "
conexao.execute sql

sql="UPDATE clc_carga INNER JOIN grades_per ON clc_carga.curso = grades_per.curso SET clc_carga.codcur = [grades_per]![codcur] WHERE clc_carga.codcur Is Null "
'conexao.execute sql
	
sql="INSERT INTO apontamento_arquivo ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase ) " & _
"SELECT '" & session("usuariomaster") & "', a.mes_base, a.chapa, c.aextra AS Evento, Sum([Extra]*60) AS ValorRef, Sum([Extra]*60) AS ValorBase " & _
"FROM qry_apontamento AS a INNER JOIN g2cursoeve c ON a.doc=c.coddoc " & _
"WHERE c.coddoc<>'' " & _
"GROUP BY a.mes_base, a.chapa, c.aextra, case when a.doc='JOR' then 'CSJ' else a.doc end " & _
"HAVING a.mes_base='" & dtaccess(mesbase) & "' AND Sum(a.Selec)<>0;"
'response.write sql & "<br><br>"
conexao.execute sql
	
sql="INSERT INTO apontamento_arquivo ( sessao, Mes_base, chapa, Evento, ValorRef, ValorBase, Ajuste ) " & _
"SELECT '" & session("usuariomaster") & "', a.Mes_base, a.chapa, g2cursoeve.falta AS Evento, " & _
"valorref=sum(case when ((case when [i] is null then 0 else i end)+(case when jd is null then 0 else jd end))=0 then null else (case when i is null then 0 else i end)+ (case when jd is null then 0 else jd end) end), " & _
"valorbase=sum(case when ((case when [i] is null then 0 else i end)+(case when jd is null then 0 else jd end))=0 then null else (case when i is null then 0 else i end)+ (case when jd is null then 0 else jd end) end), " & _
"Sum(a.Repos) AS Repos " & _
"FROM qry_apontamento AS a INNER JOIN g2cursoeve ON a.doc = g2cursoeve.coddoc " & _
"WHERE g2cursoeve.coddoc<>'' " & _
"GROUP BY a.Mes_base, a.chapa, g2cursoeve.falta, case when a.doc='JOR' then 'CSJ' else a.doc end " & _
"HAVING Mes_base='" & dtaccess(mesbase) & "' AND Sum(Selec)<>0;"
'response.write sql & "<br><br>"
conexao.execute sql

sql="INSERT INTO apontamento_arquivo ( sessao, Mes_base, chapa, Evento, ValorRef, ValorBase ) " & _
"SELECT '" & session("usuariomaster") & "', a.Mes_base, a.chapa, g2cursoeve.depen AS Evento, Sum([DP]*60) AS ValorRef, Sum([DP]*60) AS ValorBase " & _
"FROM qry_apontamento AS a INNER JOIN g2cursoeve ON a.doc = g2cursoeve.coddoc " & _
"WHERE g2cursoeve.coddoc<>'' " & _
"GROUP BY a.Mes_base, a.chapa, g2cursoeve.depen, case when a.doc='JOR' then 'CSJ' else a.doc end " & _
"HAVING Mes_base='" & dtaccess(mesbase) & "' AND Sum(Selec)<>0;"
'response.write sql & "<br><br>"
conexao.execute sql

sql="INSERT INTO apontamento_arquivo ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase ) " & _
"SELECT '" & session("usuariomaster") & "', a.mes_base, a.chapa, g2cursoeve.atraso AS Evento, Sum(a.[atraso]*1) AS ValorRef, Sum(a.[atraso]*1) AS ValorBase " & _
"FROM qry_apontamento AS a INNER JOIN g2cursoeve ON a.doc = g2cursoeve.coddoc " & _
"WHERE g2cursoeve.coddoc<>'' " & _
"GROUP BY a.mes_base, a.chapa, g2cursoeve.atraso, case when a.doc='JOR' then 'CSJ' else a.doc end " & _
"HAVING a.Mes_base='" & dtaccess(mesbase) & "' AND Sum(a.Selec)<>0;"
'response.write sql & "<br><br>"
conexao.execute sql

sql="DELETE FROM apontamento_arquivo " & _
"WHERE valorRef Is Null AND valorBase Is Null AND ajuste Is Null "
conexao.execute sql

sql="UPDATE apontamento_arquivo SET valorRef = [valorref]-[ajuste] " & _
"WHERE sessao='" & session("usuariomaster") & "' AND mes_base='" & dtaccess(mesbase) & "' AND ajuste>0; "
conexao.execute sql

sql="UPDATE apontamento_arquivo SET evento = '177', valorRef = 0 " & _
"WHERE sessao='" & session("usuariomaster") & "' AND mes_base='" & dtaccess(mesbase) & "' AND valorRef Is Null AND ajuste>0 "
conexao.execute sql
	
sql="select * from apontamento_arquivo " & _
"WHERE mes_base='" & dtaccess(mesbase) & "' and sessao='" & session("usuariomaster") & "' and valorref>0"

end if

if request.form="" then
%>
<p class=titulo>Geração de Arquivo do Apontamento dos Professores
<form method="POST" action="apont_labore.asp">
<table border="0" cellpadding="2" cellspacing="0" summary="">
<tr>
	<td class=titulo>
	<p>Mês base para emissão: <select size="1" name="mesbase">
<%
sqla="SELECT mes_base FROM clc_carga GROUP BY mes_base order by mes_base desc"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst
mesbase=rsc("mes_base")
do while not rsc.eof
%>
         <option value="<%=rsc("mes_base")%>" <%=tempt%>><%=rsc("mes_base")%></option>
<%
rsc.movenext
loop
rsc.close
%>
	</select></p>		
	</td>
</tr>
<tr>
	<td class=titulo><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></td>
</tr>
</table>
</form>
<%
else

	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="apont_prof_" & replace(request.form("mesbase"),"/","") & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	inicio=now()
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		chapa=espaco1(rs("chapa"),16)
		evento=espaco1(rs("evento"),4)
		valor=espaco1(replace(formatnumber(rs("valorref"),2),".",""),15)
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
	<td><p class=titulo>Geração de Arquivo do Apontamento dos Professores</td>
	<td><a href="../temp/<%=nomefile%>">Apontamento Prof.</a></td>
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
	response.write "<td class=campo>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"
duracao=(termino-inicio)
Response.write "<p class=realce><font size=1> Inicio: " & inicio & " Termino: " & termino & " Duracao: " & formatdatetime(duracao,3) & "</font></p>"

end if 'request.form 
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%> 
</body>
</html>