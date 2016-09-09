<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a73")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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

sql="DELETE FROM apontamento_arquivo_pos WHERE mes_base='" & dtaccess(mesbase) & "' and sessao='" & session("usuariomaster") & "' "
conexao.execute sql

sql="UPDATE clc_cargap INNER JOIN grades_per ON clc_cargap.curso = grades_per.curso SET clc_cargap.codcur = [grades_per]![codcur] WHERE clc_cargap.codcur Is Null "
'conexao.execute sql
	
'aulas
sql="INSERT INTO apontamento_arquivo_pos ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase, tipo ) " & _
"SELECT '" & session("usuariomaster") & "', a.mes_base, a.chapa, gc.sal AS Evento, Sum([aula_dada]*60) AS ValorRef, Sum([aula_dada]*60) AS ValorBase, 'A' " & _
"FROM qry_apontamentop AS a INNER JOIN g2cursoeve AS gc ON a.doc = gc.coddoc " & _
"WHERE a.aula_dada>0 GROUP BY a.mes_base, a.chapa, gc.sal " & _
"HAVING a.mes_base='" & dtaccess(mesbase) & "'"
'response.write sql & "<br><br>"
conexao.execute sql

'orientação	
sql="INSERT INTO apontamento_arquivo_pos ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase, tipo ) " & _
"SELECT '" & session("usuariomaster") & "', a.mes_base, a.chapa, gc.orient AS Evento, Sum([orientacao]*60) AS ValorRef, Sum([orientacao]*60) AS ValorBase, 'O' " & _
"FROM qry_apontamentop AS a INNER JOIN g2cursoeve AS gc ON a.doc = gc.coddoc " & _
"WHERE a.orientacao<>0 GROUP BY a.mes_base, a.chapa, gc.orient, a.doc " & _
"HAVING a.mes_base='" & dtaccess(mesbase) & "'"
'response.write sql & "<br><br>"
conexao.execute sql

'supervisao
sql="INSERT INTO apontamento_arquivo_pos ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase, tipo ) " & _
"SELECT '" & session("usuariomaster") & "', a.mes_base, a.chapa, gc.superv AS Evento, Sum([supervisao]*60) AS ValorRef, Sum([supervisao]*60) AS ValorBase, 'S' " & _
"FROM qry_apontamentop AS a INNER JOIN g2cursoeve AS gc ON a.doc = gc.coddoc " & _
"WHERE a.supervisao>0 GROUP BY a.mes_base, a.chapa, gc.superv, a.doc " & _
"HAVING a.mes_base='" & dtaccess(mesbase) & "'"
'response.write sql & "<br><br>"
conexao.execute sql

'adicional noturno
sql="INSERT INTO apontamento_arquivo_pos ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase, tipo ) " & _
"SELECT '" & session("usuariomaster") & "', a.mes_base, a.chapa, Evento=case when a.chapa='01164' then '952N' when a.chapa='01165' then '951N' when q.tipo='RT' then '387' else '387P' end, Sum([adicnot]*1.00) AS ValorRef, Sum([adicnot]*1.00) AS ValorBase, 'N' " & _
"FROM qry_apontamentop AS a left join quem_nomeacoes q on q.chapa=a.chapa collate database_default " & _
"WHERE a.adicnot>0 GROUP BY a.mes_base, a.chapa, case when a.chapa='01164' then '952N' when a.chapa='01165' then '951N' when q.tipo='RT' then '387' else '387P' end " & _
"HAVING a.mes_base='" & dtaccess(mesbase) & "' "
'response.write sql & "<br><br>"
conexao.execute sql

'aulas - evento 034/A/D
sql="INSERT INTO apontamento_arquivo_pos ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase, tipo ) " & _
"SELECT distinct '" & session("usuariomaster") & "', a.mes_base, a.chapa, '034', 0, 0, 'A' FROM qry_apontamentop a " & _
"WHERE a.aula_dada>0 GROUP BY a.mes_base, a.chapa, a.doc HAVING a.mes_base='" & dtaccess(mesbase) & "'"
'response.write sql & "<br><br>"
conexao.execute sql
sql="INSERT INTO apontamento_arquivo_pos ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase, tipo ) " & _
"SELECT distinct '" & session("usuariomaster") & "', a.mes_base, a.chapa, '034A', 0, 0, 'A' FROM qry_apontamentop a " & _
"WHERE a.aula_dada>0 GROUP BY a.mes_base, a.chapa, a.doc HAVING a.mes_base='" & dtaccess(mesbase) & "'"
'response.write sql & "<br><br>"
conexao.execute sql
sql="INSERT INTO apontamento_arquivo_pos ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase, tipo ) " & _
"SELECT distinct '" & session("usuariomaster") & "', a.mes_base, a.chapa, '034D', 0, 0, 'A' FROM qry_apontamentop a " & _
"WHERE a.aula_dada>0 GROUP BY a.mes_base, a.chapa, a.doc HAVING a.mes_base='" & dtaccess(mesbase) & "'"
'response.write sql & "<br><br>"
conexao.execute sql
sql="INSERT INTO apontamento_arquivo_pos ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase, tipo ) " & _
"SELECT distinct '" & session("usuariomaster") & "', a.mes_base, a.chapa, evento+'A', 0, 0, 'A' FROM apontamento_arquivo_pos a " & _
"WHERE a.valorref>0 and a.tipo<>'N' GROUP BY sessao, a.mes_base, a.chapa, evento HAVING a.mes_base='" & dtaccess(mesbase) & "' and sessao='" & session("usuariomaster") & "'"
'response.write sql & "<br><br>"
conexao.execute sql
sql="INSERT INTO apontamento_arquivo_pos ( sessao, mes_base, chapa, Evento, ValorRef, ValorBase, tipo ) " & _
"SELECT '" & session("usuariomaster") & "', a.mes_base, a.chapa, evento+'D', 0, 0, 'A' FROM apontamento_arquivo_pos a " & _
"WHERE a.valorref>0 and a.tipo<>'N' GROUP BY sessao, a.mes_base, a.chapa, evento HAVING a.mes_base='" & dtaccess(mesbase) & "' and sessao='" & session("usuariomaster") & "'"
'response.write sql & "<br><br>"
conexao.execute sql

sql="delete from apontamento_arquivo_pos where evento not in (Select codigo collate database_default from corporerm.dbo.pevento)"
conexao.execute sql


sql="UPDATE apontamento_arquivo_pos SET rt = 1 from apontamento_arquivo_pos a JOIN qry_rt rt ON a.chapa=rt.CHAPA collate database_default " & _
"WHERE sessao='" & session("usuariomaster") & "' AND a.mes_base='" & dtaccess(mesbase) & "'"
conexao.execute sql

sql="UPDATE apontamento_arquivo_pos SET rt = 1 " & _
"WHERE sessao='" & session("usuariomaster") & "' AND mes_base='" & dtaccess(mesbase) & "' " & _
"and chapa in ('01057','01164','01165') "
conexao.execute sql

sql="UPDATE apontamento_arquivo_pos SET rt = 0 " & _
"WHERE sessao='" & session("usuariomaster") & "' AND mes_base='" & dtaccess(mesbase) & "' " & _
"and chapa in ('00184','00266') "
'conexao.execute sql
	
sql="select * from apontamento_arquivo_pos " & _
"WHERE mes_base='" & dtaccess(mesbase) & "' and sessao='" & session("usuariomaster") & "' and (rt=0 or tipo='N') "

end if

if request.form="" then
%>
<p class=titulo>Geração de Arquivo do Apontamento dos Professores
<form method="POST" action="apont_labore_pos.asp">
<table border="0" cellpadding="2" cellspacing="0" summary="">
<tr>
	<td class=titulo>
	<p>Mês base para emissão: <select size="1" name="mesbase">
<%
sqla="SELECT mes_base FROM clc_cargap GROUP BY mes_base order by mes_base desc"
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
	nomefile="apont_pos_" & replace(request.form("mesbase"),"/","") & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
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

testepos=0
if testepos=1 then
sql1="SELECT r.CHAPA collate database_default as chapa, c.coddoc, c.CODCCUSTO, Sum([ta]*4.5) AS aulas, [JORNADA]/60 AS total, ceiling((([JORNADA]/60)/([ta]*4.5))*100)/100 AS perc " & _
"FROM (g2ch AS g INNER JOIN qry_rt AS r ON g.chapa1 = r.CHAPA collate database_default) INNER JOIN g2cursoeve AS c ON g.coddoc = c.coddoc " & _
"WHERE g.demons=1 AND '" & dtaccess(mesbase) & "' Between [inicio] And [termino] " & _
"GROUP BY r.CHAPA, c.coddoc, c.CODCCUSTO, [JORNADA]/60, ceiling((([JORNADA]/60)/([ta]*4.5))*100)/100 " & _
"union all " & _
"SELECT a.chapa, c.coddoc, c.CODCCUSTO, Sum([valorref]/60) AS aulas, [JORNADA]/60 AS total, ceiling((([JORNADA]/60)/([valorref]/60))*100)/100 AS perc " & _
"FROM qry_rt AS r INNER JOIN (g2cursoeve AS c INNER JOIN apontamento_arquivo_pos AS a ON c.sal = a.evento) ON r.CHAPA collate database_default = a.chapa " & _
"WHERE a.rt=1 AND a.sessao='" & session("usuariomaster") & "' AND a.mes_base='" & dtaccess(mesbase) & "' AND Not a.evento='034' AND Not c.CODCCUSTO=r.codsecao collate database_default " & _
"GROUP BY a.chapa, c.coddoc, c.CODCCUSTO, [JORNADA]/60, ceiling((([JORNADA]/60)/([valorref]/60))*100)/100 "
sql2="union all "
sql3="SELECT t.chapa, '' AS coddoc, r.CODSECAO collate database_default as codsecao, t.taulas, [JORNADA]/60 AS total, 100-[tperc] AS perc " & _
"FROM (select z.chapa, sum(aulas) as taulas, sum(perc) as tperc from ( " & sql1 & " ) as z group by chapa " & _
") AS t INNER JOIN qry_rt r ON t.chapa=r.CHAPA collate database_default "
sqla="select * from (" & sql1 & sql2 & sql3 & ") a order by chapa"

rs.Open sqla, ,adOpenStatic, adLockReadOnly
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

	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomerat="rateio_" & replace(request.form("mesbase"),"/","") & ".txt"
	lote=caminho & nomerat
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		chapa=espaco1(rs("chapa"),16)
		ccusto=espaco1(rs("codccusto"),25)
		valor=espaco1(replace(formatnumber(rs("perc"),2),".",""),6)
		leitura.writeline chapa & ";" & ccusto & ";" & valor & ""
	rs.movenext
	loop
	rs.close
	leitura.close
	set leitura=nothing
	set arquivo=nothing
end if 'testepos

duracao=(termino-inicio)
Response.write "<p class=realce><font size=1> Inicio: " & inicio & " Termino: " & termino & " Duracao: " & formatdatetime(duracao,3) & "</font></p>"
%>
<a href="../temp/<%=nomefile%>"><img src="../images/Diskette.gif" width="16" height="16" border="0" alt="">Arquivo Apontamento Prof.</a><br>
<!--
<a href="../temp/<%=nomerat%>"><img src="../images/Diskette.gif" width="16" height="16" border="0" alt="">Arquivo Rateio C.Custo</a><br>
-->
<%

end if 'request.form 
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%> 
</body>
</html>