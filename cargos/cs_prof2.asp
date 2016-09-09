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
<form method="POST" action="cs_prof2.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=450>
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
<td class=titulo></td>
	</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=450>
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
	h=1
	if h=1 then 
		largura=990:altura=650
	else
		largura=650:altura=990
	end if
	
if evento="0" then	
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=<%=largura%> height=<%=altura%>>
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

'if session("usuariomaster")="00259" then corta=0 else corta=1
if corta=1 then sqlrr=" and reformulacao='A' " else sqlrr=""

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
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=<%=largura%> height=<%=altura%>>
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
"GROUP BY f.dt_faixa, c.curso, c.evento, t.titulacao, t.faixa, t.titulo, t.tabela, t.reformulacao " & _
"HAVING f.dt_faixa='" & dtaccess(data) & "' and c.evento='" & rs2("evento") & "' " & _
"ORDER BY c.curso, t.reformulacao, t.titulacao, t.faixa "
sql="SELECT c.curso, t.titulacao, t.faixa, t.titulo, " & _
"ff0=max(case when nivel='0' and reformulacao='F' then t.faixasalarial else '' end),  " & _
"ff1=max(case when nivel='1' and reformulacao='F' then t.faixasalarial else '' end),  " & _
"ff2=max(case when nivel='2' and reformulacao='F' then t.faixasalarial else '' end),  " & _
"vf0=max(case when nivel='0' and reformulacao='F' then valoraula end),  " & _
"vf1=max(case when nivel='1' and reformulacao='F' then valoraula end),  " & _
"vf2=max(case when nivel='2' and reformulacao='F' then valoraula end),  " & _
"fe0=max(case when nivel='0' and reformulacao='E' then t.faixasalarial else '' end),  " & _
"fe1=max(case when nivel='1' and reformulacao='E' then t.faixasalarial else '' end),  " & _
"fe2=max(case when nivel='2' and reformulacao='E' then t.faixasalarial else '' end),  " & _
"ve0=max(case when nivel='0' and reformulacao='E' then valoraula end),  " & _
"ve1=max(case when nivel='1' and reformulacao='E' then valoraula end),  " & _
"ve2=max(case when nivel='2' and reformulacao='E' then valoraula end),  " & _
"fd0=max(case when nivel='0' and reformulacao='D' then t.faixasalarial else '' end),  " & _
"fd1=max(case when nivel='1' and reformulacao='D' then t.faixasalarial else '' end),  " & _
"fd2=max(case when nivel='2' and reformulacao='D' then t.faixasalarial else '' end),  " & _
"vd0=max(case when nivel='0' and reformulacao='D' then valoraula end),  " & _
"vd1=max(case when nivel='1' and reformulacao='D' then valoraula end),  " & _
"vd2=max(case when nivel='2' and reformulacao='D' then valoraula end),  " & _
"fb0=max(case when nivel='0' and reformulacao='B' then t.faixasalarial else '' end),  " & _
"fb1=max(case when nivel='1' and reformulacao='B' then t.faixasalarial else '' end),  " & _
"fb2=max(case when nivel='2' and reformulacao='B' then t.faixasalarial else '' end),  " & _
"vb0=max(case when nivel='0' and reformulacao='B' then valoraula end),  " & _
"vb1=max(case when nivel='1' and reformulacao='B' then valoraula end),  " & _
"vb2=max(case when nivel='2' and reformulacao='B' then valoraula end),  " & _
"fa0=max(case when nivel='0' and reformulacao='A' then t.faixasalarial else '' end),  " & _
"fa1=max(case when nivel='1' and reformulacao='A' then t.faixasalarial else '' end),  " & _
"fa2=max(case when nivel='2' and reformulacao='A' then t.faixasalarial else '' end),  " & _
"va0=max(case when nivel='0' and reformulacao='A' then valoraula end),  " & _
"va1=max(case when nivel='1' and reformulacao='A' then valoraula end),  " & _
"va2=max(case when nivel='2' and reformulacao='A' then valoraula end)  " & _
"FROM csd_cursos AS c INNER JOIN csd_titulos AS t ON c.tabela = t.tabela  " & _
"INNER JOIN csd_faixas AS f ON t.faixasalarial = f.faixasalarial " & _
"WHERE '" & dtaccess(data) & "' Between [ivigencia] And [fvigencia] " & sqlrr & " and t.titulacao not in ('C','E','GE')  " & _
"and f.dt_faixa='" & dtaccess(data) & "' and c.evento='" & rs2("evento") & "' and t.reformulacao<>'C'  " & _
"GROUP BY c.curso, t.titulacao, t.faixa, t.titulo  ORDER BY c.curso, t.titulacao, t.faixa "
'response.write sql
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=<%=largura-50%>>
<tr><td class=titulop align="center" colspan=2 rowspan=2>&nbsp;Titulação</td>
	<td class=titulop align="center" colspan=9>Faixas</td>
</tr>
<tr>
<%for a=65 to 7%>
<td class=titulop align="center"><%=chr(a)%></td>
<%next%>
</tr>
</tr>
<%
rs.movefirst
do while not rs.eof
vf0=rs("vf0")
vf1=rs("vf1")

ve0=rs("ve0"):if ve0<vf0 then ve0=vf0
ve1=rs("ve1"):if ve1<vf1 then ve1=vf1
've2=rs("ve2")
vd0=rs("vd0"):if vd0<ve0 then vd0=ve0
vd1=rs("vd1"):if vd1<ve1 then vd1=ve1
'vd2=rs("vd2"):if vd2<ve2 then vd2=ve2
vb0=rs("vb0"):if vb0<vd0 then vb0=vd0
vb1=rs("vb1"):if vb1<vd1 then vb1=vd1
'vb2=rs("vb2"):if vb2<vd2 then vb2=vd2
va0=rs("va0"):if va0<vb0 then va0=vb0
va1=rs("va1"):if va1<vb1 then va1=vb1
va2=rs("va2"):if va2<vb2 then va2=vb2
if vf0=vf1 then vf0="" else vf0=formatnumber(vf0,2)
if vf1=ve0 then vf1="" else vf1=formatnumber(vf1,2)
if ve0=ve1 then ve0="" else ve0=formatnumber(ve0,2)
if ve1=vd0 then ve1="" else ve1=formatnumber(ve1,2)
if vd0=vd1 then vd0="" else vd0=formatnumber(vd0,2)
if vd1=vb0 then vd1="" else vd1=formatnumber(vd1,2)
if vb0=vb1 then vb0="" else vb0=formatnumber(vb0,2)
if vb1=va0 then vb1="" else vb1=formatnumber(vb1,2)
if va0=va1 then va0="" else va0=formatnumber(va0,2)
if va1=va2 then va1="" else va1=formatnumber(va1,2)
va2=formatnumber(va2,2)

%>
<tr>
<td class="campop">&nbsp;<%=rs("faixa")%></td>
<td class="campop" align="left">&nbsp;<%=rs("titulo")%></td>
<td class="campop" align="center"><b><%=vf1%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("ff1")%></td>
<td class="campop" align="center"><b><%=ve1%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("fe1")%></td>
<td class="campop" align="center"><b><%=vd0%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("fd0")%></td>
<td class="campop" align="center"><b><%=vd1%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("fd1")%></td>
<td class="campop" align="center"><b><%=vb0%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("fb0")%></td>
<td class="campop" align="center"><b><%=vb1%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("fb1")%></td>
<td class="campop" align="center"><b><%=va0%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("fa0")%></td>
<td class="campop" align="center"><b><%=va1%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("fa1")%></td>
<td class="campop" align="center"><b><%=va2%>&nbsp;</b><p style="margin-top:0;margin-bottom:0;color:Gray;font-size:7pt;text-align:right"><%=rs("fa2")%></td>

</tr>
<%
rs.movenext
loop
rs.close
%>
</table>
<br>
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