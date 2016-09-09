<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a34")="N" or session("a34")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Grades - Quadro de Horário do Professor</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<!--<script language="JavaScript" type="text/javascript" src="../date.js"></script> -->
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
Sub CentralizaCelulas
	leitura.writeline "objXL.Selection.HorizontalAlignment = -4108" 'xlCenter"
	leitura.writeline "objXL.Selection.VerticalAlignment = -4107" 'xlBottom"
	leitura.writeline "objXL.Selection.WrapText = False"
	leitura.writeline "objXL.Selection.Orientation = 0"
	leitura.writeline "objXL.Selection.AddIndent = False"
	leitura.writeline "objXL.Selection.ShrinkToFit = False"
	leitura.writeline "objXL.Selection.MergeCells = False"
	'leitura.writeline "objXL.Selection.Merge"
End Sub

Sub Esquerda
	leitura.writeline "objXL.Selection.HorizontalAlignment = -4131" 'xlLeft
	leitura.writeline "objXL.Selection.VerticalAlignment = -4107" 'xlBottom
	leitura.writeline "objXL.Selection.WrapText = False"
	leitura.writeline "objXL.Selection.Orientation = 0"
	leitura.writeline "objXL.Selection.AddIndent = False"
	leitura.writeline "objXL.Selection.IndentLevel = 0"
	leitura.writeline "objXL.Selection.ShrinkToFit = False"
	leitura.writeline "objXL.Selection.MergeCells = True"
end sub

Sub Bordas
	leitura.writeline "objXL.Selection.Borders(5).LineStyle = xlNone"
	leitura.writeline "objXL.Selection.Borders(6).LineStyle = xlNone"
	leitura.writeline "objXL.Selection.Borders(7).LineStyle = xlContinuous"
	leitura.writeline "objXL.Selection.Borders(7).Color = RGB(0,0,0)"
	leitura.writeline "objXL.Selection.Borders(8).LineStyle = xlContinuous"
	leitura.writeline "objXL.Selection.Borders(8).Color = RGB(255,255,255)"
	leitura.writeline "objXL.Selection.Borders(9).LineStyle = xlContinuous"
	leitura.writeline "objXL.Selection.Borders(9).Color = 0"
	leitura.writeline "objXL.Selection.Borders(10).LineStyle = xlContinuous"
	leitura.writeline "objXL.Selection.Borders(10).Color = 0"
	leitura.writeline "objXL.Selection.Borders(11).LineStyle = xlContinuous"
	leitura.writeline "objXL.Selection.Borders(11).Color = 0"
	leitura.writeline "objXL.Selection.Borders(12).LineStyle = xlContinuous"
End Sub

dim conexao, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form<>"" then
	if request.form("B3")<>"" then
		finaliza=1
	else
		finaliza=0
	end if
end if

if finaliza=0 then
%>
<p class=titulo>Seleção para impressão de Quadro de Horário do Professor</p>
<form method="POST" action="grd_professor.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Nome do professor</td>
</tr>
<tr>
	<td class=titulo><select size="1" name="chapa">
		<option value="0" selected>Selecione um professor</option>
<%
sqla="SELECT f.CHAPA, F.NOME FROM g2ch g INNER JOIN " & _
"grades_aux_prof as F ON g.chapa1 = F.CHAPA " & _
"WHERE g.coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') " & _
"GROUP BY F.CHAPA, F.NOME ORDER BY F.NOME "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
	<option <%if request.form("chapa")=rs("chapa") then response.write "selected"%> value="<%=rs("chapa")%>"><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
if request.form("database")<>"" then database=request.form("database") else database=now()
if request.form("database")<>"" then database=request.form("database") else database=dateserial(year(now),month(now)+1,1)
database=formatdatetime(database,2)
%>  
	<option value="02270">ROSA</option>
	<option value="00610">Helio</option>
	<option value="00315">Claudio Aguiar</option>
		</select>
	</td>
</tr>
<tr>
	<td class=titulo>Horário vigente na data de &nbsp;<input type=text name="database" value=<%=database%> size=10 class=a >
	<!-- onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')"--> 
	&nbsp;<input type="checkbox" name="naluno" value="ON">Imprimir nº alunos</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="400">
<tr><td align="center" class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3"></td></tr>
</table>
</form>
<hr>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="400">
<tr><td align="left" class=campo>
1-há uma quebra de página entre cada período letivo/turno.<br>
2-configure seu navegador para imprimir no modo paisagem (<i>landscape</i>).<br>
3-demora em média 30 segundos para ser gerado.
</td></tr>
</table>

<%
end if 'finaliza 0

'******************************** inicio impressao
if finaliza=1 then
	dim disciplina(15), planilha(15), planilha2(15), planilha3(15), backfundo(15), compl(15)
	for t1=0 to 15
		for t2=0 to 0
			disciplina(t1)=t1*t2+rnd(3)
		next
	next
	linha=1:coluna=1
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="quadro_professor_" & session("usuariomaster") & ".vbs"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	leitura.writeline "Dim objXL"
	leitura.writeline "Set objXL = WScript.CreateObject(""Excel.Application"")"
	leitura.writeline "objXL.Visible = TRUE"
	leitura.writeline "objXL.WorkBooks.Add"
	leitura.writeline "objXL.Cells.Select"
	leitura.writeline "objXL.Selection.Font.Name = ""Tahoma"""
	leitura.writeline "objXL.Selection.Font.Size = 8"
	leitura.writeline "objXL.Cells.EntireColumn.AutoFit"
    leitura.writeline "objXL.Selection.VerticalAlignment = -4160" 'xlTop"

	numero=0
	chapa=request.form("chapa")
	database=request.form("database")
	datarel=dtaccess(database)

	sql0="select nome from corporerm.dbo.pfunc where chapa='" & chapa & "' "
	sql0="select nome from grades_aux_prof where chapa='" & chapa & "' "
	rs.Open sql0, ,adOpenStatic, adLockReadOnly
	nomechapa=rs("nome")
	rs.close
	sessao=session("usuariomaster") & "P"
	sql2="delete from g2temp where sessao='" & sessao & "'":conexao.execute sql2
	sql3="insert into g2temp (sessao, descricao, horini, horfim) select '" & sessao & "', descricao, horini, horfim from g2defhor " & _
	"where codtn=" & turno & " and tipocurso=" & tipocurso & " group by descricao, horini, horfim"
	sql3="insert into g2temp (sessao, descricao, horini, horfim) select '" & sessao & "', descricao, horini, horfim from g2ch g " & _
	"where g.chapa1='" & chapa & "' and '" & datarel & "' between inicio and termino group by descricao, horini, horfim "
	conexao.execute sql3
	for a=2 to 7
		sql4="update g2temp set [" & a & "]=g.codhor from g2temp t, g2ch g " & _
		"where t.horini=g.horini and t.horfim=g.horfim and g.diasem=" & a & " and chapa1='" & chapa & "' and '" & datarel & "' between inicio and termino and sessao='" & sessao & "'"
		conexao.execute sql4
	next
'sql="select * from g2temp where sessao='" & sessao & "' "
'rs1.Open sql, ,adOpenStatic, adLockReadOnly
'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs1.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs1.fields(a).name) & "</td>"
'next
'response.write "</tr>"
'do while not rs1.eof 
'response.write "<tr>"
'for a= 0 to rs1.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs1.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs1.movenext:loop
'response.write "</table>"
'response.write "# " & rs1.recordcount & "<br>"
'rs1.close
'*************** fim teste **********************

	numero=0
	'****** dia da semana
	numero=0
	for z=2 to 7
		redim preserve diasemana(numero)
		diasemana(numero)=z
		numero=numero+1
	next
					
coluna=1
linha=linha+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & chapa & " - " & nomechapa & """":linha=linha+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & "Grade Horária em " & formatdatetime(database,2) & """":linha=linha+1

leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = ""Horário"""
for tabela=0 to ubound(diasemana)
	coluna=coluna+1
	leitura.writeline "objXL.Columns(" & coluna & ").ColumnWidth = 30"
	leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & weekdayname(diasemana(tabela)) & """"
next
letra=chr(65+coluna)
leitura.writeline "objXL.Range(""A" & linha & ":" & letra & linha & """).Select"
leitura.writeline "objXL.Selection.Font.Bold = True"
CentralizaCelulas
leitura.writeline "objXL.Range(""A1"").Select"
linha=linha+1
					
%>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=980>
<tr>
	<td class=grupo colspan="<%=3+ubound(diasemana)+1%>"><b><%=chapa & " - " & nomechapa %></b></td>
</tr>
<tr>
	<td class="campol" colspan="<%=1+ubound(diasemana)+1%>">Grade Horária em <%=formatdatetime(database,2)%></td>
</tr>
<tr>
	<td class="campot"r>Horário</td>
	<%for tabela=0 to ubound(diasemana)%>
	<td class="campot"r align="center"><%=weekdayname(diasemana(tabela))%>
	</td>
	<%next%>
</tr>
<tr>
<%

sqlq="select * from g2temp where sessao='" & sessao & "' order by horini "
rs.Open sqlq, ,adOpenStatic, adLockReadOnly

linhainicio=linha
coluna=1
	do while not rs.eof			
	'for d=0 to ubound(horaini)
%>
	<td class="campoa"r nowrap><%=rs("descricao")%></td>
<%
coluna=1
'linha=linhainicio
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & rs("descricao") & """"

for e=0 to ubound(diasemana)
	celula1="1":celula3="1"
	chapa=chapa
	diasem=diasemana(e)
	if isnull(rs.fields(e+4)) then codhor=0 else codhor=rs.fields(e+4)
	sql6="select g.chapa1, g.coddoc, c.curso, g.perlet, g.codsala, g.codmat, m.materia, g.codtur, g.inicio, g.termino, juntar, dividir, jturma, dturma from g2ch g, corporerm.dbo.umaterias m, g2cursoeve c " & _
	"where c.coddoc=g.coddoc and g.codmat=m.codmat collate database_default and g.codhor=" & codhor & " and diasem=" & diasem & " and '" & datarel & "' between inicio and termino and chapa1='" & chapa & "' "
	'response.write "<br><br>" &sql6
	rs1.Open sql6, ,adOpenStatic, adLockReadOnly
	celula1=""
	if rs1.recordcount>0 then
		do while not rs1.eof
		codcur=rs1("coddoc")
		curso=rs1("curso")
		codmat=rs1("codmat")
		materia=rs1("materia")
		chapa1=rs1("chapa1")
		perlet=rs1("perlet")
		codtur=rs1("codtur")
		sala=rs1("codsala"):if sala="0" then sala=""
if request.form("naluno")="ON" then
	sql3="SELECT Count(umc.MATALUNO) AS alunos " & _
	"FROM corporerm.dbo.USITMAT AS usm1 INNER JOIN ((((corporerm.dbo.UMATRICPL umc INNER JOIN corporerm.dbo.UMATALUN uma ON (umc.GRADE=uma.GRADE) AND (umc.CODPER=uma.CODPER) AND (umc.CODCUR=uma.CODCUR) AND (umc.PERLETIVO=uma.PERLETIVO) AND (umc.MATALUNO=uma.MATALUNO) AND (umc.CODCOLIGADA=uma.CODCOLIGADA) AND (umc.CODFILIAL=uma.CODFILIAL)) " & _
	"INNER JOIN corporerm.dbo.UMATERIAS um ON (uma.CODCOLIGADA=um.CODCOLIGADA) AND (uma.CODMAT=um.CODMAT)) " & _
	"INNER JOIN corporerm.dbo.UALUCURSO a ON (umc.GRADE=a.GRADE) AND (umc.CODPER=a.CODPER) AND (umc.CODCUR=a.CODCUR) AND (umc.MATALUNO=a.MATALUNO) AND (umc.CODCOLIGADA=a.CODCOLIGADA)) " & _
	"INNER JOIN corporerm.dbo.USITMAT usm ON a.STATUS=usm.CODSITMAT) ON usm1.CODSITMAT=uma.STATUS " & _
	"WHERE umc.PERLETIVO='" & rs1("perlet") & "' AND CODTUR='" & rs1("codtur") & "' AND uma.CODMAT='" & rs1("codmat") & "' " & _
	"and uma.STATUS In ('01','07','08','09','10','18','19','20','47','48','46','70') "
	rs2.Open sql3, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then alunos="Alunos: " & rs2("alunos") & "<br>" else alunos="erro " & rs1("codtur")
	rs2.close
else
	alunos=""
end if
if isnull(rs1("codsala")) then sala="" else sala="Sala: " & rs1("codsala") & "<br>"
if rs1("juntar")=1 then
	obs="<font color=blue>Junta turma " & rs1("jturma") & "<br>"
elseif rs1("dividir")=1 then
	obs="<font color=red>Divide turma" & "<br>"
else
	obs=""
end if
compl(e)=alunos & obs
		celula1=celula1 & "<b>" & curso & " / " & codtur & "</b><br>" & materia & "<br>Sala " & sala
		celula2=curso & " / " & codtur
		celula2b=materia & duplicado
		celula2c=sala
		rs1.movenext
		celula1=celula1 & "<br>"
		loop
		fundo="#FFFFFF"
	else
		celula1="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		celula2=""
		celula2b=""
		celula2c=""
		fundo="#CCCCCC"
		compl(e)=""
	end if
	rs1.close
	disciplina(e)=celula1
	planilha(e)=celula2
	planilha2(e)=celula2b
	planilha3(e)=celula2c
	backfundo(e)=fundo
next 'dia semana e

for e=0 to ubound(diasemana)
linhas=1
%>					
	<td class="campor" valign=top style="background-color: <%=backfundo(e)%>" rowspan="<%=linhas%>"><%=disciplina(e)%><%=compl(e)%>
	</td>
<%
coluna=coluna+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & planilha(e) & """ & chr(10) & """ & planilha2(e) & """"
	if planilha3(e)<>"" or planilha(e)<>"0" then
		leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = objXL.Cells(" & linha & ", " & coluna & ").Value & chr(10) & """ & planilha3(e) & """"
	end if
next 'sturma e
linha=linha+1
%>
</tr>
<%
	rs.movenext: loop
	'next 'horainicio d
	rs.close
%>
</table>
<br>
<%
leitura.close
set leitura=nothing
set arquivo=nothing
%>
<a href="../temp/<%=nomefile%>"><img src="../images/MSExcel.gif" width="16" height="16" border="0" alt=""></a>
<%
end if 'finaliza 1

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>