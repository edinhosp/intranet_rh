<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Termos de Abertura e Encerramento</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:40px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
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
<p class=titulo>Seleção para impressão dos Termos de Abertura e Encerramento do livro de ponto</p>
<form method="POST" action="ponto_termos.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=450>
<tr><td class=titulo colspan=2>Curso:</td>
	</tr>
<tr><td class=titulo colspan=2><select size="1" name="codcur" class=a>
	<option value="0" selected>Selecione um curso</option>
<%
'sqla="SELECT codcur, curso from grades_2 where codcur in (select codcur from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY codcur, curso order by curso "
sqla="SELECT gr.coddoc, gr.CURSO " & _
"FROM grades_5 AS gr " & _
"WHERE gr.coddoc In (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') " & _
"AND gr.inicio>(getdate()-365) " & _
"GROUP BY gr.coddoc, gr.CURSO ORDER BY gr.CURSO; "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof 
%>
<option <%if codcur=rs("coddoc") then response.write "selected "%> value="<%=rs("coddoc")%>"><%=rs("curso")%></option>
<%
rs.movenext
loop
rs.close
%>  
	</select></td>
</tr>

<tr><td class=titulo>Grade:</td>
	<td class=titulo>Período:</td>
	</tr>
<tr>
	<td class=titulo><select size="1" name="grade">
		<option value="" selected>Selecione uma grade</option>
		<option value="Anual">Anual</option>
		<option value="Semestral">Semestral</option>
		</select></td>	
	<td class=titulo><select size="1" name="periodo">
		<option value="" selected>Selecione um período</option>
		<option value="Matutino">Matutino</option>
		<option value="Vespertino">Vespertino</option>
		<option value="Noturno">Noturno</option>
		<option value="Integral">Integral</option>
		</select></td>	
</tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=450>
<tr>
	<td class=titulo valign=top width=150>Data da Abertura: </td>
	<td class=titulo valign=top>
	<input type="text" name="dt_abertura" size="12" maxlength="12">
	</tr>
</tr>
<tr>
	<td class=titulo valign=top>Data do Encerramento: </td>
	<td class=titulo valign=top>
	<input type="text" name="dt_encerramento" size="12" maxlength="12">
	</td>
</tr>
<tr>
	<td class=titulo valign=top>Páginas utilizadas: </td>
	<td class=titulo valign=top>
	<input type="text" name="paginas" size="5" maxlength="5" class=a>
	</td>
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
	codcur=request.form("codcur")
	sql="select curso from g2cursoeve where coddoc='" & codcur & "'"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	curso=rs("curso")
	rs.close
	grade=request.form("grade")
	periodo=request.form("periodo")
	dt_abertura=request.form("dt_abertura")
	dt_encerramento=request.form("dt_encerramento")
	dtabertura=day(dt_abertura) & " de " & monthname(month(dt_abertura)) & " de " & year(dt_abertura)
	dtencerramento=day(dt_encerramento) & " de " & monthname(month(dt_encerramento)) & " de " & year(dt_encerramento)
	paginas=request.form("paginas")
	if paginas<>"" then
		if paginas>1 then paginas=paginas+1
		if paginas>1 then paginan=trim(extenson(paginas))
	else
		paginas="________"
		paginan="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		paginan=paginan&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	end if
	if grade<>"" then gradet=" - Grade " & grade else gradet=""
	if periodo<>"" then periodot=" - Período: " & periodo else periodot=""

tamanho=640
%>
<!-- borda -->
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho+10%>" height=990>
<tr><td class=campo valign="center" height=15 align="right">
<b><font size=5>1</b>
</td></tr>
<tr><td class=campo valign="center" height=100% align="center">
<!-- ponto -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="450">
<tr>
	<td class="campop" align="center"><b><p style="margin-top:5;margin-bottom:5;color:Black;font-size:22pt;text-align:center;">
	TERMO DE ABERTURA</b>
	<br><br><br></td>
</tr>
<tr>
	<td class="campop" align="left"><p style="margin-top:5;margin-bottom:5;color:Black;font-size:16pt;text-align:justify;">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Este livro, que contém <%=paginas%> (<%=paginan%>) folhas numeradas tipograficamente, por mim rubricadas, 
	destina-se a Registro de Presença do Corpo Docente do Curso de <%=curso%><%=gradet%><%=periodot%>.
	<br><br><br></td>
</tr>
<tr>
	<td class="campop" align="left"><p style="margin-top:5;margin-bottom:5;color:Black;font-size:14pt;text-align:left;">
	Osasco, <%=dtabertura%>
	<br><br><br></td>
</tr>
<tr>
	<td class="campop" align="left" style="border-bottom: 2 solid #000000">
	<br><br><br></td>
</tr>
</table>
<!-- ponto -->
<!-- borda -->
</td></tr>
</table>

<DIV style="page-break-after:always"></DIV>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho+10%>" height=990>
<tr><td class=campo valign="center" height=15 align="right">
<b><font size=5><%=paginas%></b>
</td></tr>
<tr><td class=campo valign="center" height=100% align="center">
<!-- ponto -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="450">
<tr>
	<td class="campop" align="center"><b><p style="margin-top:5;margin-bottom:5;color:Black;font-size:22pt;text-align:center;">
	TERMO DE ENCERRAMENTO</b>
	<br><br><br></td>
</tr>
<tr>
	<td class="campop" align="left"><p style="margin-top:5;margin-bottom:5;color:Black;font-size:16pt;text-align:justify;">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Contém este livro <%=paginas%> (<%=paginan%>) folhas numeradas tipograficamente, por mim rubricadas, 
	e destina-se ao uso indicado no Termo de Abertura.
	<br><br><br></td>
</tr>
<tr>
	<td class="campop" align="left"><p style="margin-top:5;margin-bottom:5;color:Black;font-size:14pt;text-align:left;">
	Osasco, <%=dtencerramento%>
	<br><br><br></td>
</tr>
<tr>
	<td class="campop" align="left" style="border-bottom: 2 solid #000000">
	<br><br><br></td>
</tr>
</table>
<!-- ponto -->
<!-- borda -->
</td></tr>
</table>

<%
set rs2=nothing
end if ' finaliza=1

'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>