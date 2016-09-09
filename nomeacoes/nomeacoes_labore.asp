<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a28")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Geração de Arquivo Nomeações</title>
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
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

sqlh="select valor from corporerm.dbo.pvalfix where codigo='Nom' and getdate() between iniciovigencia and finalvigencia"
rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then hora=rs2("valor") else hora=0
rs2.close

if request.form<>"" then
	ano=request.form("ano")
	mes=request.form("mes")
	mesq=numzero(mes,2)
	mandfim=dateserial(year(now),12,31)
	udia=day(dateserial(ano,mes+1,1)-1)
	conexao.execute "delete from ttcodigofixo"
	
sqlh="select valor from corporerm.dbo.pvalfix where codigo='Nom' and '" & dtaccess(dateserial(ano,mes+1,1)-1) & "' between iniciovigencia and finalvigencia"
rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then hora=rs2("valor") else hora=0
rs2.close
	
if mes=1 then
	sql1="SELECT n.CHAPA, n.NOME, n.CODEVE, c.descricao, Sum(n.CH) AS TotalCH, c.fator, 30 as prop, (Sum([ch]*[fator])/" & udia & ")*(case when month(mand_ini)=" & mes & " and year(mand_ini)=" & ano & " then " & udia & "-day(mand_ini)+1 else case when month(mand_fim)=" & mes & " and year(mand_fim)=" & ano & " then day(mand_fim) else " & udia & " end end) AS horas, " & _
	"ref=(sum(ch*fator)/" & udia & ")*(case when month(mand_ini)=" & mes & " and year(mand_ini)=" & ano & " then " & udia & "-day(mand_ini)+1 else case when month(mand_fim)=" & mes & " and year(mand_fim)=" & ano & " then day(mand_fim) else " & udia & " end end)*case when n.codeve in ('173') then " & nraccess(hora) & " else 60 end * (case when n.janeiro=1 then 0 else 1 end) " & _
	"FROM n_indicacoes AS n INNER JOIN cnv_atividade AS c ON (c.id_nomeacao=n.id_nomeacao) AND (n.CODEVE=c.codevento) " & _
	"WHERE (cast('" & mes & "'+'/'+'01'+'/'+'" & ano & "' as datetime) between mand_ini and (case when mand_fim is null then '" & dtaccess(mandfim) & "' else mand_fim end)) " & _
	"OR (cast('" & mes & "'+'/'+'" & udia & "'+'/'+'" & ano & "' as datetime) between mand_ini and (case when mand_fim is null then '" & dtaccess(mandfim) & "' else mand_fim end)) " & _
	"GROUP BY n.CHAPA, n.NOME, n.CODEVE, c.descricao, c.fator, n.janeiro, n.MAND_INI, n.MAND_FIM " & _
	"HAVING not n.CODEVE is null "
else
	sql1="SELECT n.CHAPA, n.NOME, n.CODEVE, c.descricao, Sum(n.CH) AS TotalCH, c.fator, 30 as prop, (Sum([ch]*[fator])/" & udia & ")*(case when month(mand_ini)=" & mes & " and year(mand_ini)=" & ano & " then " & udia & "-day(mand_ini)+1 else case when month(mand_fim)=" & mes & " and year(mand_fim)=" & ano & " then day(mand_fim) else " & udia & " end end) AS horas, " & _
	"ref=(sum(ch*fator)/" & udia & ")*(case when month(mand_ini)=" & mes & " and year(mand_ini)=" & ano & " then " & udia & "-day(mand_ini)+1 else case when month(mand_fim)=" & mes & " and year(mand_fim)=" & ano & " then day(mand_fim) else " & udia & " end end)*case when n.codeve in ('173') then " & nraccess(hora) & " else 60 end " & _
	"FROM n_indicacoes AS n INNER JOIN cnv_atividade AS c ON (c.id_nomeacao=n.id_nomeacao) AND (n.CODEVE=c.codevento) " & _
	"WHERE (cast('" & mes & "'+'/'+'01'+'/'+'" & ano & "' as datetime) between mand_ini and (case when mand_fim is null then '" & dtaccess(mandfim) & "' else mand_fim end)) " & _
	"OR (cast('" & mes & "'+'/'+'" & udia & "'+'/'+'" & ano & "' as datetime) between mand_ini and (case when mand_fim is null then '" & dtaccess(mandfim) & "' else mand_fim end)) " & _
	"GROUP BY n.CHAPA, n.NOME, n.CODEVE, c.descricao, c.fator, n.MAND_INI, n.MAND_FIM " & _
	"HAVING not n.CODEVE is null "
end if
	sql2="SELECT n.CHAPA, n.NOME, n.CODEVE, c.descricao, Sum(n.complemento) AS TotalCH, c.fator, " & udia & " AS prop, " & _
	"Sum(n.complemento) as horas, Sum(n.complemento)*(case when n.codeve in ('173') then " & nraccess(hora) & " else 60 end) AS ref " & _
	"FROM n_indicacoes AS n INNER JOIN cnv_atividade AS c ON (c.id_nomeacao=n.id_nomeacao) " & _
	"AND (n.CODEVE=c.codevento) WHERE FPG_INI=cast('" & mes & "'+'/'+'01'+'/'+'" & ano & "' as datetime) " & _
	"GROUP BY n.CHAPA, n.NOME, n.CODEVE, c.descricao, c.fator, n.MAND_INI, n.MAND_FIM " & _
	"HAVING not n.CODEVE is null " 

	sql3="SELECT n.CHAPA, n.NOME, c.cod_acumulador, 'Acumulador' AS descr, Sum(0) AS ch1, 0 AS fat, 30 AS prop, 0 AS horas, 0 AS ref " & _
	"FROM n_indicacoes AS n INNER JOIN cnv_atividade AS c ON n.CODEVE=c.codevento " & _
	"WHERE (cast('" & mes & "'+'/'+'01'+'/'+'" & ano & "' as datetime) between mand_ini and (case when mand_fim is null then '" & dtaccess(mandfim) & "' else mand_fim end)) " & _
	"OR (cast('" & mes & "'+'/'+'" & udia & "'+'/'+'" & ano & "' as datetime) between mand_ini and (case when mand_fim is null then '" & dtaccess(mandfim) & "' else mand_fim end)) " & _
	"GROUP BY n.CHAPA, n.NOME, c.cod_acumulador  " & _
	"HAVING not c.cod_acumulador is null "

	sql0=sql1 & " union all " & sql2 & " union all " & sql3
	
	sql="SELECT t.CHAPA, t.NOME, CODEVE, descricao, Sum(horas) AS horas, Sum(ref) AS ref " & _
	"FROM (" & sql0 & ") as t " & _
	"inner join corporerm.dbo.pfunc f on f.chapa=t.chapa collate database_default " & _
	"where (t.chapa not in (select chapa collate database_Default from qry_rt) and codeve<>'108') " & _
	" and f.codsituacao in ('A','F','Z') " & _
	"GROUP BY t.CHAPA, t.NOME, CODEVE, descricao " & _
	"ORDER BY CODEVE, t.CHAPA "
end if

if request.form="" then 
%>
<p class=titulo>Geração de arquivo das Nomeações para o RM Labore
<form method="POST" action="nomeacoes_labore.asp" name="form">
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
	<td class=fundo>Hora atividade</td>
	<td class=fundo><%=hora%></td>
</tr>
<tr>
	<td class=titulo colspan="2"><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></td>
</tr>
<tr></tr>
<tr><td class="campor" colspan=2>Mês de geração é igual ao mês competência</td></tr>
</table>
</form>
<%
else
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="nomeacoes" & request.form("ano") & mesq & ".txt"
	lote=caminho & nomefile
	response.write lote
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
	'rs.close
	termino=now()
	duracao=(termino-inicio)
	'Response.write "Inicio: " & inicio & "<br>Termino: " & termino & "<br>Duracao: " & formatdatetime(duracao,3)
	leitura.close
	set leitura=nothing
	set arquivo=nothing

%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width='690'>
<tr>
	<td class=campo><p class=titulo>Geração de arquivo das Nomeações para o RM Labore</td>
	<td class="campor">Hora atividade: <%=hora%></td>
	<td class=campo><a href="../temp/<%=nomefile%>"><img src="../images/Diskette.gif" border="0" width="16" height="16" alt=""></a></td>
</tr>
</table>
<%
'rs.Open sql, ,adOpenStatic, adLockReadOnly
total=0
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse' width='690'>"
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
%>
<%
end if 'request.form

%> 
<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>