<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a78")="N" or session("a78")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cargos e Salários - Professores</title>
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
mesagora=month(now)
anoagora=year(now)
diaagora=day(now)

if request.form<>"" then
	if request.form("B3")<>"" then
		finaliza=1
	else
		finaliza=0
	end if
end if

if finaliza=0 then
%>
<p class=titulo>Seleção para impressão de Tabela Salarial</p>
<form method="POST" action="cs_prof.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=550>
<tr><td class=titulo>Data da Tabela:</td>
<td class=titulo>Tabela:</td>
<td class=titulo>Agrupar?</td></tr>
<tr><td class=titulo><select size="1" name="data">
<%
sqla="SELECT data FROM csd_obs ORDER by data desc"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof 
%>
<option value="<%=rs("data")%>"><%=rs("data")%></option>
<%
rs.movenext
loop
rs.close
%>  
	</select></td>
<td class=titulo><select size="1" name="setor">
	<option value="0" selected>Todos Cursos</option>
	<option value="T">Apenas Tabela de Faixas</option>
	<option value="V">Apenas Tabela Variável</option>
<%
sqla="SELECT evento, curso FROM csd_cursos GROUP BY evento, curso ORDER BY curso "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof 
%>
<option value="<%=rs("evento")%>"><%=rs("curso")%></option>
<%
rs.movenext
loop
rs.close
%>  
	</select></td>	
<td class=titulo><input type="checkbox" name="agrupar" value="ON">
<br>Ult.Recl.? <input type="checkbox" name="cortar" value="ON">
</td>
	</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=550>
<tr><td align="center" class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3"></td></tr>
</table>
</form>
<hr>
<%
end if 'finaliza=0

'******************************** inicio impressao
if finaliza=1 then
	data=request.form("data")
	evento=request.form("setor")
	sql="select observacao from csd_obs where data='" & dtaccess(data) & "' "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	observacao=rs("observacao")
	rs.close

if evento="0" then	
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650 height=990>
<tr>
	<td class=campo height=100% colspan=2 align="center" valign="center" style="border: 1px solid #000000">
	<font size=5>TABELA SALARIAL<br>PROFESSORES<br><%=monthname(month(data)) & "/" & year(data)%></font></td>
</tr>
<tr><td class=campo height=20 colspan=2 style="border-bottom: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">
	<%=observacao%></td></tr>
<tr><td class=campo height=20><b>Recursos Humanos&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%pagina=pagina+1%></td>
	<td class=campo align="right"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr></table>
<%
end if 'setor=0

if evento="T" or evento="0" then	
	if evento="0" then response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650 height=990>
<tr><td class=campo height=20 style="border-bottom: 1px solid #000000"><b>CENTRO UNIVERSITÁRIO FIEO</td>
	<td class=campo align="right" style="border-bottom: 1px solid #000000"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr>
<tr>
	<td class=campo height=100% colspan=2 align="center" valign=top>
<font size=2><b><br>Tabela de Faixas Salariais<br>Início Vigência: <%=monthname(month(data)) & "/" & year(data)%><br>&nbsp;</font>
<!-- quadro do setor -->
<%
sql="SELECT dt_faixa, faixa, faixasalarial, titulacao, valoraula, g, [e], m, d " & _
"FROM csd_faixas WHERE dt_faixa='" & dtaccess(data) & "' " & _
"ORDER BY ordem, dt_faixa, left(faixasalarial,1), faixa "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<%
inicio=0
rs.movefirst
do while not rs.eof
%>
<%
if inicio=0 or linhaf>50 then
if inicio=1 then response.write "</table>":response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=600>
<tr><td class=titulo>&nbsp;Nível Salarial</td>
<td class=titulo align="center" width=100>Graduado</td>
<td class=titulo align="center" width=100>Especialista</td>
<td class=titulo align="center" width=100>Mestre</td>
<td class=titulo align="center" width=100>Doutor</td></tr>
<%
linhaf=2
pagina=pagina+1
end if
%>

<tr>
<td class=campo align="center">&nbsp;<%=rs("faixasalarial")%></td>
<% if rs("g")=true then %>
<td class=campo align="right"><%=formatnumber(rs("valoraula"),2)%>&nbsp;</td>
<% else
response.write "<td class=titulo align=""right"">&nbsp;</td>"
end if%>
<% if rs("e")=true then %>
<td class=campo align="right"><%=formatnumber(rs("valoraula"),2)%>&nbsp;</td>
<% else
response.write "<td class=titulo align=""right"">&nbsp;</td>"
end if%>
<% if rs("m")=true then %>
<td class=campo align="right"><%=formatnumber(rs("valoraula"),2)%>&nbsp;</td>
<% else
response.write "<td class=titulo align=""right"">&nbsp;</td>"
end if%>
<% if rs("d")=true then %>
<td class=campo align="right"><%=formatnumber(rs("valoraula"),2)%>&nbsp;</td>
<% else
response.write "<td class=titulo align=""right"">&nbsp;</td>"
end if%>
</td>
<%
linhaf=linhaf+1 : inicio=1
rs.movenext
loop
rs.close
%>
</table>
<!-- fim quadro do setor -->
	</td>
</tr>
<tr><td class=campo height=20 colspan=2 style="border-bottom: 1px solid #000000">
	<%=observacao%></td></tr>
<tr><td class=campo height=20><b>Recursos Humanos&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pág. <%response.write pagina%></td>
	<td class=campo align="right"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr></table>
<%
end if 'setor=0 or setor=T

if evento="V" or evento="0" then	
	if evento="0" then response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650 height=990>
<tr><td class=campo height=20 style="border-bottom: 1px solid #000000"><b>CENTRO UNIVERSITÁRIO FIEO</td>
	<td class=campo align="right" style="border-bottom: 1px solid #000000"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr>
<tr>
	<td class=campo height=100% colspan=2 align="center" valign=top>
<font size=2><b><br>Tabela Salarial Variável<br>Início Vigência: <%=monthname(month(data)) & "/" & year(data)%><br>&nbsp;</font>
<!-- quadro do setor -->
<%
sql="SELECT dt_extra, tipo, ordem, codigo, verba, observacao, valor_hora " & _
"FROM csd_extras WHERE dt_extra='" & dtaccess(data) & "' and tipo='Variável' " & _
"ORDER BY tipo DESC , ordem "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=600>
<tr><td class=titulo>&nbsp;Cod.</td>
<td class=titulo align="center">Descrição</td>
<td class=titulo align="center">Observação</td>
<td class=titulo align="center">Valor</td>
<%
rs.movefirst:do while not rs.eof
if isnull(rs("valor_hora")) then valor_hora="&nbsp;" else valor_hora=formatnumber(rs("valor_hora"),2)
%>
<tr>
<td class=campo align="center">&nbsp;<%=rs("codigo")%></td>
<td class=campo><%=rs("verba")%>&nbsp;</td>
<td class=campo><%=rs("observacao")%>&nbsp;</td>
<td class=campo align="right"><%=valor_hora%>&nbsp;</td>
</td>
<%
rs.movenext:loop
%>
</table>
<%
end if
rs.close
%>

<font size=2><b><br>Tabela Salarial Pós-Graduação<br>Início Vigência: <%=monthname(month(data)) & "/" & year(data)%><br>&nbsp;</font>
<%
sql="SELECT dt_extra, tipo, ordem, codigo, verba, observacao, valor_hora, reformulacao " & _
"FROM csd_extras WHERE dt_extra='" & dtaccess(data) & "' and tipo='Pós-Graduação' " & _
"ORDER BY reformulacao, tipo DESC , ordem "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=600>
<tr><td class=titulo>&nbsp;Cod.</td>
<td class=titulo align="center">Descrição</td>
<td class=titulo align="center">Observação</td>
<td class=titulo align="center">Valor</td>
<%
rs.movefirst:do while not rs.eof
if rs("reformulacao")<>lastreforma then 
	if rs("reformulacao")="A" then texto="Admissões até 31/jan/2005"
	if rs("reformulacao")="B" then texto="Admissões a partir de 01/fev/2005"
	if rs("reformulacao")="C" then texto="Admissões a partir de 01/fev/2007"
	if rs("reformulacao")="D" then texto="Admissões a partir de 01/ago/2009"
	if rs("reformulacao")="E" then texto="Admissões a partir de 01/fev/2010"
	if rs("reformulacao")="F" then texto="Admissões a partir de 01/jan/2016"
	if rs("reformulacao")="G" then texto="Admissões a partir de 01/jul/2016"
	if rs.recordcount<=4 then texto="---"
	response.write "<tr><td class=""campop"" style='border:2px solid #000000' colspan=5><b>" & texto & "</td></tr>"
end if
if isnull(rs("valor_hora")) then valor_hora="&nbsp;" else valor_hora=formatnumber(rs("valor_hora"),2)
%>
<tr>
<td class=campo align="center">&nbsp;<%=rs("codigo")%></td>
<td class=campo><%=rs("verba")%>&nbsp;</td>
<td class=campo><%=rs("observacao")%>&nbsp;</td>
<td class=campo align="right"><%=valor_hora%>&nbsp;</td>
</td>
<%
lastreforma=rs("reformulacao")
rs.movenext:loop
%>
</table>
<%
end if
rs.close
%>

<!-- fim quadro do setor -->
	</td>
</tr>
<tr><td class=campo height=20 colspan=2 style="border-bottom: 1px solid #000000">
	<%=observacao%></td></tr>
<tr><td class=campo height=20><b>Recursos Humanos&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pág. <%pagina=pagina+1:response.write pagina%></td>
	<td class=campo align="right"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr></table>
<%
end if 'setor=0 or setor=T

if session("usuariomaster")="00259" then corta=0 else corta=0
if request.form("cortar")="ON" then corta=1 else corta=0
if corta=1 then sqlrr=" and reformulacao='A' " else sqlrr=""
if request.form("agrupar")="" then

if evento<>"T" and evento<>"V" then
	if evento="0" then sqls="" else sqls=" AND evento='" & evento & "' "
sql2="SELECT curso, evento, tabela FROM csd_cursos " & _
"WHERE '" & dtaccess(data) & "' Between [ivigencia] And [fvigencia] and tabela<>0 " & sqls & _
"GROUP BY evento, curso, tabela " & _
"ORDER BY curso, tabela "
'response.write sql2
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if rs2.recordcount>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650 height=990>
<tr><td class=campo height=20 style="border-bottom: 1px solid #000000"><b>CENTRO UNIVERSITÁRIO FIEO</td>
	<td class=campo align="right" style="border-bottom: 1px solid #000000"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr>
<tr>
	<td class=campo height=100% colspan=2 align="center" valign=top>
<font size=2><b><br>Tabela Salarial de Valor Aula<br><%=monthname(month(data)) & "/" & year(data)%><br>&nbsp;</font>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=95%>
<tr><td class="campop">Curso: <b><%=rs2("curso")%><br></td></tr>
<tr><td class="campop">Evento: <b><%=rs2("evento")%><br></td></tr>
</table>
<br>
<!-- quadro do setor -->
<%
sql="SELECT f.dt_faixa, c.curso, c.evento, t.titulacao, t.faixa, t.titulo, t.tabela, t.reformulacao, " & _
"faixa0=max(case when nivel='0' then t.faixasalarial else '' end), faixa1=max(case when nivel='1' then t.faixasalarial else '' end), " & _
"faixa2=max(case when nivel='2' then t.faixasalarial else '' end), " & _
"valor0=max(case when nivel='0' then valoraula end), valor1=max(case when nivel='1' then valoraula end), " & _
"valor2=max(case when nivel='2' then valoraula end) " & _
"FROM csd_cursos AS c INNER JOIN (csd_titulos AS t INNER JOIN csd_faixas AS f ON t.faixasalarial = f.faixasalarial) ON c.tabela = t.tabela " & _
"WHERE '" & dtaccess(data) & "' Between [ivigencia] And [fvigencia] " & sqlrr &  _
"and t.titulacao not in ('C','E','GE') " & _
"and t.printable=1 " & _
"GROUP BY f.dt_faixa, c.curso, c.evento, t.titulacao, t.faixa, t.titulo, t.tabela, t.reformulacao " & _
"HAVING f.dt_faixa='" & dtaccess(data) & "' and c.evento='" & rs2("evento") & "' " & _
"ORDER BY c.curso, t.reformulacao, t.titulacao, t.faixa "
'response.write sql
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=600>
<tr><td class=titulop align="center" colspan=2>&nbsp;Titulação</td>
<td class=titulop align="center" width=100>Nível<br>Admissional</td>
<td class=titulop align="center" width=100>Nível<br>Pós-Experiência</td>
<td class=titulop align="center" width=100>Nível<br>Especial</td>
</tr>
<%
rs.movefirst
do while not rs.eof
if rs("reformulacao")<>lastreforma then 
	if rs("reformulacao")="A" then texto="Admissões até 31/jan/2005"
	if rs("reformulacao")="B" then texto="Admissões a partir de 01/fev/2005"
	if rs("reformulacao")="C" then texto="Admissões a partir de 01/fev/2007"
	if rs("reformulacao")="D" then texto="Admissões a partir de 01/ago/2009"
	if rs("reformulacao")="E" then texto="Admissões a partir de 01/fev/2010"
	if rs("reformulacao")="F" then texto="Admissões a partir de 01/jan/2016"
	if rs("reformulacao")="G" then texto="Admissões a partir de 01/jul/2016"

	if corta=1 then texto="-"

	response.write "<tr><td class=""campop"" style='border:2px solid #000000' colspan=5><b>" & texto & "</td></tr>"
end if
if isnull(rs("valor0")) then valor0="&nbsp;" else valor0=formatnumber(rs("valor0"),2)
if isnull(rs("valor1")) then valor1="&nbsp;" else valor1=formatnumber(rs("valor1"),2)
if isnull(rs("valor2")) then valor2="&nbsp;" else valor2=formatnumber(rs("valor2"),2)
if rs("titulacao")="C" or rs("titulacao")="E" then aviso="<i><b>(1)</b></i>" else aviso=""
%>
<tr>
<td class="campop">&nbsp;<%=rs("faixa")%></td>
<td class="campop" align="left">&nbsp;<%=rs("titulo")%>&nbsp;<%=aviso%></td>
<td class="campop" align="center"><b><%=valor0%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("faixa0")%></td>
<td class="campop" align="center"><b><%=valor1%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("faixa1")%></td>
<td class="campop" align="center"><b><%=valor2%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("faixa2")%></td>
</tr>
<%
lastreforma=rs("reformulacao")
rs.movenext
loop
rs.close
%>
</table>
<br>
<%if corta<>1 then%>
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:8pt;text-align:left"><i>(1)</i> Faixas extintas conforme artigo 12 da Resolução nº 20 de 25/11/2003
<%end if%>
<br>
<!-- fim quadro do setor -->
	</td>
</tr>

<tr><td class=campo height=20 colspan=2 style="border-bottom: 1px solid #000000">
	<%=observacao%></td></tr>
<tr><td class=campo height=20><b>Recursos Humanos&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pág. <%pagina=pagina+1:response.write pagina%></td>
	<td class=campo align="right"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr></table>
<%
rs2.movenext
loop
rs2.close
end if 'setor <> T/V

else 'request.form agrupar

if evento<>"T" and evento<>"V" then
	if evento="0" then sqls="" else sqls=" AND evento='" & evento & "' "
sql2="SELECT tabela FROM csd_cursos " & _
"WHERE tabela<>0 and '" & dtaccess(data) & "' Between [ivigencia] And [fvigencia] " & sqls & _
"GROUP BY tabela "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
if rs2.recordcount>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650 height=990>
<tr><td class=campo height=20 style="border-bottom: 1px solid #000000"><b>CENTRO UNIVERSITÁRIO FIEO</td>
	<td class=campo align="right" style="border-bottom: 1px solid #000000"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr>
<tr>
	<td class=campo height=100% colspan=2 align="center" valign=top>
<font size=2><b><br>Tabela Salarial de Valor Aula<br><%=monthname(month(data)) & "/" & year(data)%><br>&nbsp;</font>
<br>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=600>
<tr><td class=titulop>Cursos abrangidos por esta tabela</td></tr>
<tr><td class=titulop>Curso</td></tr>
<%
sql="SELECT evento, curso FROM csd_cursos " & _
"WHERE #" & dtaccess(data) & "# Between [ivigencia] And [fvigencia] and tabela=" & rs2("tabela") & " " & _
"GROUP BY evento, curso ORDER BY curso "
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
%>
<tr><td class="campop"><%=rs("evento")%> - <b><%=rs("curso")%></td></tr>
<%
rs.movenext
loop
rs.close
%>
</table>
<br>
<!-- quadro do setor -->
<%
sql="SELECT f.dt_faixa, t.tabela, t.titulacao, t.faixa, t.titulo, t.reformulacao, Max(IIf([nivel]='0',[t]![faixasalarial],'')) AS faixa0, Max(IIf([nivel]='1',[t]![faixasalarial],'')) AS faixa1, Max(IIf([nivel]='2',[t]![faixasalarial],'')) AS faixa2, Max(IIf([nivel]='0',[valoraula])) AS valor0, Max(IIf([nivel]='1',[valoraula])) AS valor1, Max(IIf([nivel]='2',[valoraula])) AS valor2 " & _
"FROM csd_cursos AS c INNER JOIN (csd_titulos AS t INNER JOIN csd_faixas AS f ON t.faixasalarial = f.faixasalarial) ON c.tabela = t.tabela " & _
"WHERE #" & dtaccess(data) & "# Between [ivigencia] And [fvigencia] " & _
"GROUP BY f.dt_faixa, t.tabela, t.titulacao, t.faixa, t.titulo, t.reformulacao " & _
"HAVING f.dt_faixa=#" & dtaccess(data) & "# and t.tabela=" & rs2("tabela") & " " & _
"ORDER BY t.tabela, t.reformulacao, t.titulacao, t.faixa "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=600>
<tr><td class=titulop align="center" colspan=2>&nbsp;Titulação</td>
<td class=titulop align="center" width=100>Nível<br>Admissional</td>
<td class=titulop align="center" width=100>Nível<br>Pós-Experiência</td>
<td class=titulop align="center" width=100>Nível<br>Especial</td>
</tr>
<%
rs.movefirst
do while not rs.eof
if rs("reformulacao")<>lastreforma then 
	if rs("reformulacao")="A" then texto="Admissões até 31/jan/2005"
	if rs("reformulacao")="B" then texto="Admissões a partir de 01/fev/2005"
	if rs("reformulacao")="C" then texto="Admissões a partir de 01/fev/2007"
	if rs("reformulacao")="D" then texto="Admissões a partir de 01/ago/2009"
	if rs("reformulacao")="E" then texto="Admissões a partir de 01/fev/2010"
	response.write "<tr><td class=""campop"" style='border:2px solid #000000' colspan=5><b>" & texto & "</td></tr>"
end if
if isnull(rs("valor0")) then valor0="&nbsp;" else valor0=formatnumber(rs("valor0"),2)
if isnull(rs("valor1")) then valor1="&nbsp;" else valor1=formatnumber(rs("valor1"),2)
if isnull(rs("valor2")) then valor2="&nbsp;" else valor2=formatnumber(rs("valor2"),2)
if rs("titulacao")="C" or rs("titulacao")="E" then aviso="<i><b>(1)</b></i>" else aviso=""
%>
<tr>
<td class="campop">&nbsp;<%=rs("faixa")%></td>
<td class="campop" align="left">&nbsp;<%=rs("titulo")%>&nbsp;<%=aviso%></td>
<td class="campop" align="center"><b><%=valor0%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("faixa0")%></td>
<td class="campop" align="center"><b><%=valor1%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("faixa1")%></td>
<td class="campop" align="center"><b><%=valor2%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("faixa2")%></td>
</tr>
<%
lastreforma=rs("reformulacao")
rs.movenext
loop
rs.close
%>
</table>
<br>
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:8pt;text-align:left"><i>(1)</i> Faixas extintas conforme artigo 12 da Resolução nº 20 de 25/11/2003
<br>
<!-- fim quadro do setor -->
	</td>
</tr>

<tr><td class=campo height=20 colspan=2 style="border-bottom: 1px solid #000000">
	<%=observacao%></td></tr>
<tr><td class=campo height=20><b>Recursos Humanos&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pág. <%pagina=pagina+1:response.write pagina%></td>
	<td class=campo align="right"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr></table>
<%
rs2.movenext
loop
end if 'recordcount
rs2.close
end if 'setor <> T/V

end if 'request.form agrupar
%>

<%
end if 'finaliza=1
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
'SELECT [cs_areas].[id_setor], [cs_areas].[setor], [cs_carreira].[imagem], [cs_cargos].[ordem], [cs_cargos].[cargo], [cs_salarios].[data], [cs_salarios].[n1], [cs_salarios].[n2], [cs_salarios].[n3], [cs_salarios].[n4], [cs_salarios].[n5], [cs_obs].[observacao] FROM (((cs_cargos INNER JOIN cs_areas ON [cs_cargos].[id_setor]=[cs_areas].[id_setor]) INNER JOIN cs_carreira ON [cs_cargos].[id_setor]=[cs_carreira].[id_setor]) INNER JOIN cs_salarios ON ([cs_cargos].[id_setor]=[cs_salarios].[id_setor]) AND ([cs_cargos].[id_cargo]=[cs_salarios].[id_cargo])) INNER JOIN cs_obs ON [cs_salarios].[data]=[cs_obs].[data] WHERE ((([cs_salarios].[data])=[forms]![cs_administrativos]![cmbdata])) ORDER BY [cs_cargos].[ordem]; 
%>
</body>
</html>