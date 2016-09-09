<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Aniversariantes do Mês</title>
<link rel="stylesheet" type="text/css" href="../../<%=session("estilo")%>">
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
	sql="select f.chapa, f.codsituacao, Month(p.dtnascimento) AS Mes, Day(p.dtnascimento) AS Dia, f.nome, f.codsecao, campus=case substring(f.codsecao,1,2) when '01' then 'Narciso' when '02' then 'Brás' when '03' then 'V.Yara' when '04' then 'Jd.Wilson' else '-' end, p.dtnascimento, s.descricao AS Setor, corpo=case f.codtipo when 'T' then 'Estagiário' else case f.codsindicato when '03' then 'Professor' else 'Administrativo' end end " & _
	"from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.psecao s " & _
	"where f.codpessoa=p.codigo and f.codsecao=s.codigo and (f.chapa<'10000' or f.chapa>='90000') and  f.chapa not in ('00322','01094') and f.codsituacao<>'D' " & _
	"and month(p.dtnascimento)=" & mes & " " & _
	"order by month(p.dtnascimento), day(p.dtnascimento) "
end if
if request("mes")<>"" then
	ano=year(now)
	mes=request("mes")
	mesq=numzero(mes,2)
	udia=day(dateserial(ano,mes+1,1)-1)
	sql="select f.chapa, f.codsituacao, Month(p.dtnascimento) AS Mes, Day(p.dtnascimento) AS Dia, f.nome, f.codsecao, campus=case substring(f.codsecao,1,2) when '01' then 'Narciso' when '02' then 'Brás' when '03' then 'V.Yara' when '04' then 'Jd.Wilson' else '-' end, p.dtnascimento, s.descricao AS Setor, corpo=case f.codtipo when 'T' then 'Estagiário' else case f.codsindicato when '03' then 'Professor' else 'Administrativo' end end " & _
	"from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.psecao s " & _
	"where f.codpessoa=p.codigo and f.codsecao=s.codigo and (f.chapa<'10000' or f.chapa>='90000') and  f.chapa not in ('00322','01094') and f.codsituacao<>'D' " & _
	"and month(p.dtnascimento)=" & mes & " " & _
	"order by month(p.dtnascimento), day(p.dtnascimento) "
end if

if request.form="" and request("mes")="" then
%>
<p class=titulo>Lista de aniversariantes do mês
<form method="POST" action="isabel.asp" name="form">
	<p>Ano <input type="text" name="ano" size="6" value="<%=year(now())%>"> 
	Mês <input type="text" name="mes" size="4" value="<%=month(now())%>"></p>
	<p><input type="submit" value="Visualizar" name="Gerar" class="button"></p>
</form>
<%
else
%>
<table border="0" cellpadding="3" width="650" cellspacing="0" style="border-collapse:collapse">
<tr>
	<td class=campo colspan="2"><img border="0" src="../images/logo_centro_universitario_unifieo_big.jpg" width=225></td>
	<td class=campo colspan="3" align="center"><b>ANIVERSARIANTES DO MÊS DE &nbsp;<%=ucase(monthname(mes))%></b></td>
</tr>
</table>

<table border="1" cellpadding="3" width="650" cellspacing="0" style="border-collapse:collapse">
<tr>
	<td class=titulo>Dia</td>
	<td class=titulo>Nome</td>
	<td class=titulo></td>
	<td class=titulo>Campus</td>
	<td class=titulo>Seção</td>
</tr>
<%
rs.Open sql, ,adOpenStatic, adLockReadOnly
linha=3
rs.movefirst
do while not rs.eof
if linha>47 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p style='margin-top: 0; margin-bottom: 0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border=0 cellpadding=3 width=650 cellspacing=0 style='border-collapse:collapse'>"
	response.write "<tr>"
	response.write "<td class=campo colspan='2'><img border='0' src='../images/logo_centro_universitario_unifieo_big.jpg' width=225></td>"
	response.write "<td class=campo colspan='3' align='center'><b>ANIVERSARIANTES DO MÊS DE &nbsp;" & ucase(monthname(mes)) & "</b></td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border=1 cellpadding=3 width=650 cellspacing=0 style='border-collapse:collapse'>"
	response.write "<tr>"
	response.write "<td class=titulo>Dia</td>"
	response.write "<td class=titulo>Nome</td>"
	response.write "<td class=titulo></td>"
	response.write "<td class=titulo>Campus</td>"
	response.write "<td class=titulo>Seção</td>"
	response.write "</tr>"
	linha=3
end if
%>
<tr>
	<td class=campo align="center"><%=rs("dia")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("corpo")%></td>
	<td class=campo><%=rs("campus")%></td>
	<td class=campo><%=rs("setor")%></td>
</tr>
<%
linha=linha+1
rs.movenext
loop
rs.close
linha=linha+1
pagina=pagina+1
%>
</table>
<%
response.write "<p style='margin-top: 0; margin-bottom: 0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
end if
%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>