<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a77")="N" or session("a77")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cargos e Salários - Administrativos</title>
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
<form method="POST" action="cs_adm.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=250>
<tr><td class=titulo>Data da Tabela:</td>
<td class=titulo>Tabela:</td>
</tr>
<tr><td class=titulo><select size="1" name="data">
<%
sqla="SELECT data FROM cs_obs order by data desc"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
<option value="<%=rs("data")%>"><%=rs("data")%></option>
<%
rs.movenext:loop
rs.close
%>  
	</select></td>
<td class=titulo><select size="1" name="setor">
	<option value="0" selected>Todos Setores</option>
<%
sqla="SELECT id_setor, setor from cs_areas order by setor "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof 
%>
<option value="<%=rs("id_setor")%>"><%=rs("setor")%></option>
<%
rs.movenext
loop
rs.close
%>  
	</select></td>	
	</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=250>
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
	setor=request.form("setor")
	sql="select observacao from cs_obs where data='" & dtaccess(data) & "' "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	observacao=rs("observacao")
	rs.close

if setor="0" then	
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650 height=990>
<tr>
	<td class=campo height=100% colspan=2 align="center" valign="center" style="border: 1px solid #000000">
	<font size=5>TABELA SALARIAL<br>ADMINISTRATIVOS<br><%=monthname(month(data)) & "/" & year(data)%></font></td>
</tr>
<tr><td class=campo height=20 colspan=2 style="border-bottom: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">
	<%=observacao%></td></tr>
<tr><td class=campo height=20><b>Recursos Humanos&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%pagina=pagina+1%></td>
	<td class=campo align="right"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr></table>
<%
end if 'setor=0

if setor="0" then sqls="" else sqls=" AND cs_salarios.id_setor='" & setor & "' "
sql2="SELECT cs_salarios.id_setor, cs_areas.setor, cs_salarios.data " & _
"FROM cs_salarios INNER JOIN cs_areas ON cs_salarios.id_setor = cs_areas.id_setor " & _
"GROUP BY cs_salarios.id_setor, cs_areas.setor, cs_salarios.data " & _
"HAVING cs_salarios.data='" & dtaccess(data) & "' AND cs_salarios.id_setor<>'NCLA' " & sqls & _
"ORDER BY cs_areas.setor "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if rs2.recordcount>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650 height=990>
<tr><td class=campo height=20 style="border-bottom: 1px solid #000000"><b>Pró-Reitoria Administrativa</td>
	<td class=campo align="right" style="border-bottom: 1px solid #000000"><b><font color="#0000ff">UNI</font><font color="#ff0000">FIEO</font></td>
</tr>
<tr>
	<td class=campo height=100% colspan=2 align="center" valign=top>
<font size=3><b><br>Plano de Carreira<br><%=rs2("setor")%><br><%=monthname(month(data)) & "/" & year(data)%><br>&nbsp;</font>
<!-- quadro do setor -->
<%
sql="SELECT a.id_setor, a.setor, c.ordem, c.cargo, s.data, s.n1, s.n2, s.n3, s.n4, s.n5, c.horas " & _
"FROM (cs_cargos AS c INNER JOIN cs_areas AS a ON c.id_setor = a.id_setor) " & _
"INNER JOIN cs_salarios AS s ON (c.id_cargo = s.id_cargo) AND (c.id_setor = s.id_setor) " & _
"WHERE s.data='" & dtaccess(data) & "' and a.id_setor='" & rs2("id_setor") & "' " & _
"ORDER BY a.setor, c.ordem "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=600>
<tr><td class=titulo width=200>&nbsp;Cargo</td><td class=titulo align="center">Horas</td>
<td class=titulo align="center">N1</td><td class=titulo align="center">N2</td>
<td class=titulo align="center">N3</td><td class=titulo align="center">N4</td>
<td class=titulo align="center">N5</td></tr>
<%
rs.movefirst
do while not rs.eof
%>
<tr>
<td class=campo>&nbsp;<%=rs("cargo")%></td>
<td class=campo align="right"><%=rs("horas")%>&nbsp;</td>
<td class=campo align="right"><%=formatnumber(rs("n1"),2)%>&nbsp;</td>
<td class=campo align="right"><%=formatnumber(rs("n2"),2)%>&nbsp;</td>
<td class=campo align="right"><%=formatnumber(rs("n3"),2)%>&nbsp;</td>
<td class=campo align="right"><%=formatnumber(rs("n4"),2)%>&nbsp;</td>
<td class=campo align="right"><%=formatnumber(rs("n5"),2)%>&nbsp;</td>
</tr>
<%
lastsetor=rs("setor")
rs.movenext
loop
rs.close
%>
</table>
<br>
<br>
<img border="0" src="../carreira/<%=rs2("id_setor")%>.jpg" width="630">

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