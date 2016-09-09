<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a33")="N" or session("a33")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Grades - Quadro de Horário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<!-- <script language="JavaScript" type="text/javascript" src="../date.js"></script> -->
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

inicio=now()

dim conexao, chapach, rs, rs1, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	if request.form("B3")<>"" then
		finaliza=1
	else
		finaliza=0
	end if
end if

if finaliza=0 then
%>
<p class=titulo>Seleção para impressão de Quadro de Horário</p>
<form method="POST" action="grd_quadro.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Escolha o Curso Graduação</td></tr>
<tr>
	<td class=titulo><select size="1" name="codcur">
	<option value="0" selected>Selecione um curso</option>
<%
sqla="SELECT gc.coddoc, gc.CURSO " & _
"FROM g2ch gr INNER JOIN g2cursoeve gc ON gr.coddoc=gc.coddoc " & _
"GROUP BY gc.coddoc, gc.CURSO ORDER BY gc.CURSO; "
if session("usuariomaster") then
end if
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof 
%>
<option <%if request.form("coddoc")=rs("coddoc") then response.write "selected "%> value="<%=rs("coddoc")%>"><%=rs("curso")%></option>
<%
rs.movenext
loop
rs.close
if request.form("database")<>"" then database=request.form("database") else database=dateserial(year(now),month(now)+1,1)
database=formatdatetime(database,2)
%>  
	</select></td></tr>
<tr>
	<td class=titulo>Horário vigente na data de &nbsp;<input type=text name="database" value=<%=database%> size=10 class=a >
	<!--onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')"-->
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
	dim disciplina(20), planilha(20), planilha2(20), planilha3(20), backfundo(20), compl(20)
	for t1=0 to 20
		for t2=0 to 0
			disciplina(t1)=t1*t2+rnd(3)
		next
	next
	linha=1:coluna=1
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="quadro_horario_" & session("usuariomaster") & ".vbs"
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
	codcur=request.form("codcur")
	database=request.form("database")
	datarel=dtaccess(database)

	sql1="SELECT c.coddoc, c.curso, g.perlet FROM g2ch g, g2cursoeve c " & _
	"WHERE c.coddoc=g.coddoc AND g.deletada=0 AND '" & datarel & "' Between inicio And termino " & _
	"GROUP BY c.coddoc, c.curso, g.perlet HAVING c.coddoc='" & codcur & "' ORDER BY g.perlet "
	'response.write "<br>Etapa 1: " & sql1               '****************************
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	'response.write "<br>Resul 1: " & rs.recordcount     '****************************
	rs.movefirst
	do while not rs.eof
		redim preserve periodoletivo(numero)
		periodoletivo(numero)=rs("perlet")
		curso=rs("curso")
	rs.movenext
	numero=numero+1
	loop
	rs.close

	for a=0 to ubound(periodoletivo)
		numero=0
		sql2="SELECT turno, descturno FROM g2ch g, eturnos e " & _
		"WHERE deletada=0 and g.turno=e.codturno AND g.coddoc='" & codcur & "' AND perlet Like '" & periodoletivo(a) & "' AND '" & datarel & "' between inicio and termino " & _
		"GROUP BY turno, descturno ORDER BY turno "
		'response.write "<br>Etapa 2: " & sql2            '****************************
		rs.Open sql2, ,adOpenStatic, adLockReadOnly
		'response.write "<br>Resul 2: " & rs.recordcount  '****************************
		rs.movefirst
		do while not rs.eof
			redim preserve turnograde(numero)
			redim preserve turnonome(numero)
			turnograde(numero)=rs("turno")
			turnonome(numero)=rs("descturno")
		rs.movenext
		numero=numero+1
		loop
		rs.close

		for b=0 to ubound(turnograde)
			numero=0
			sql5="SELECT serie, turma, codtur FROM g2ch g " & _
			"WHERE deletada=0 AND g.coddoc='" & codcur & "' AND perlet Like '" & periodoletivo(a) & "%' AND '" & datarel & "' between inicio and termino " & _
			"and turno=" & turnograde(b) & " GROUP BY serie, turma, codtur "
			'response.write "<br>Etapa 3: " & sql5              '****************************
			rs.Open sql5, ,adOpenStatic, adLockReadOnly
			'response.write "<br>Resul 3: " & rs.recordcount    '****************************
			rs.movefirst
			do while not rs.eof
				redim preserve sturma(numero)
				redim preserve serie(numero)
				sturma(numero)=rs("codtur"):'response.write sturma(numero) & " "
				serie(numero)=rs("serie")
			rs.movenext
			numero=numero+1
			loop
			rs.close

coluna=1
linha=linha+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & curso & """"
linha=linha+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = ""Período Letivo: " & periodoletivo(a) & """"
linha=linha+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = ""Período: " & turnonome(b) & """"
linha=linha+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = ""Dia"""
coluna=coluna+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = ""Horário"""
for tabela=0 to ubound(sturma)
	coluna=coluna+1
	leitura.writeline "objXL.Columns(" & coluna & ").ColumnWidth = 40"
	leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & sturma(tabela) & """"
next
letra=chr(65+coluna)
leitura.writeline "objXL.Range(""A" & linha & ":" & letra & linha & """).Select"
leitura.writeline "objXL.Selection.Font.Bold = True"
CentralizaCelulas
leitura.writeline "objXL.Range(""A1"").Select"
linha=linha+1
					
%>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=grupo colspan="<%=3+ubound(sturma)+1%>"><b><%=curso%></b> - em <%=formatdatetime(database,2)%></td>
</tr>
<tr>
	<td class="campol"><%=periodoletivo(a)%></td>
	<td class="campol" colspan="<%=1+ubound(sturma)+1%>"><%=turnonome(b)%></td>
</tr>
<tr>
	<td class="campot"r>Dia</td>
	<td class="campot"r>Horário</td>
	<%for tabela=0 to ubound(sturma)%>
	<td class="campot"r align="center"><%=sturma(tabela)%>
	</td>
	<%next%>
</tr>
<%
			idiasem=0
			sqlg="select top 100 percent turno, diasem, descricao " 
			for c=0 to ubound(sturma)
				sqlg=sqlg & ",'" & sturma(c) & "'=max(case when serie=" & serie(c) & " then codhor else null end) "
			next ' for c
			sqlg=sqlg & "from g2ch " & _
			"where coddoc='" & codcur & "' and '" & datarel & "' between inicio and termino and turno=" & turnograde(b) & " and perlet='" & periodoletivo(a) & "' " & _
			"group by turno, diasem, descricao order by turno, diasem, descricao "
			'response.write "<br>Etapa 4: " & sqlg              '****************************
			rs.Open sqlg, ,adOpenStatic, adLockReadOnly
			'response.write "<br>Resul 4: " & rs.recordcount    '****************************
			do while not rs.eof
%>
<tr>
<%
			turmas=rs.fields.count-1
			if ultimodia<>rs("diasem") then idiasem=0
			if idiasem=0 then
				sqldiasem="select count(diasem) as linhas from (" & sqlg & ") z where diasem=" & rs("diasem")
				rs1.Open sqldiasem, ,adOpenStatic, adLockReadOnly : linhasdiasem=rs1("linhas") : rs1.close
				idiasem=1
%>
	<td class="campoa"r rowspan="<%=linhasdiasem%>" style="border-top:2px solid #000000" ><%=weekdayname(rs("diasem"))%></td>
<%
			end if

linhainicio=linha
coluna=1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & weekdayname(rs("diasem")) & """"

			for t=0 to turmas-3
				celula1="1":celula3="1"
				if rs.fields(t+3)>0 then   'troquei de grades_aux_prof para grades_aux_prof2 para mostrar demitidos nos quadros de horarios antigos
				sql6="select g.chapa1, f.nome, g.codmat, m.materia, juntar, jturma, dividir, dturma, codsala, inicio, termino, g.chapa2, codtur " & _
				"from g2chi g, grades_aux_prof2 f, corporerm.dbo.umaterias m " & _
				"where coddoc='" & codcur & "' and '" & datarel & "' between inicio and termino and turno=" & turnograde(b) & " and perlet='" & periodoletivo(a) & "' " & _
				"and diasem=" & rs("diasem") & " and codhor=" & rs.fields(t+3) & " and codtur='" & rs.fields(t+3).name & "' " & _
				"and g.chapa1=f.chapa and g.codmat=m.codmat collate database_default "
'				else
'				sql6="select null chapa1, null nome, null codmat, null materia, 0 juntar, null jturma, 0 dividir, null dturma, null codsala, null inicio, null termino, null chapa2, null codtur "
'				end if
				'response.write "<br>Etapa 6: " & sql6 & "<br>"
				rs1.Open sql6, ,adOpenStatic, adLockReadOnly
				if rs1.recordcount>0 then
				celula1="":celula2="":celula2b="":celula2c=""						
				rs1.movefirst
				do while not rs1.eof
				'*******alunos******
	if request.form("naluno")="ON" then
		sql3="SELECT Count(umc.MATALUNO) AS alunos " & _
		"FROM corporerm.dbo.USITMAT AS usm1 INNER JOIN ((((corporerm.dbo.UMATRICPL umc INNER JOIN corporerm.dbo.UMATALUN uma ON (umc.GRADE=uma.GRADE) AND (umc.CODPER=uma.CODPER) AND (umc.CODCUR=uma.CODCUR) AND (umc.PERLETIVO=uma.PERLETIVO) AND (umc.MATALUNO=uma.MATALUNO) AND (umc.CODCOLIGADA=uma.CODCOLIGADA) AND (umc.CODFILIAL=uma.CODFILIAL)) " & _
		"INNER JOIN corporerm.dbo.UMATERIAS um ON (uma.CODCOLIGADA=um.CODCOLIGADA) AND (uma.CODMAT=um.CODMAT)) " & _
		"INNER JOIN corporerm.dbo.UALUCURSO a ON (umc.GRADE=a.GRADE) AND (umc.CODPER=a.CODPER) AND (umc.CODCUR=a.CODCUR) AND (umc.MATALUNO=a.MATALUNO) AND (umc.CODCOLIGADA=a.CODCOLIGADA)) " & _
		"INNER JOIN corporerm.dbo.USITMAT usm ON a.STATUS=usm.CODSITMAT) ON usm1.CODSITMAT=uma.STATUS " & _
		"WHERE umc.PERLETIVO='" & periodoletivo(a) & "' AND CODTUR='" & rs1("codtur") & "' AND uma.CODMAT='" & rs1("codmat") & "' " & _
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
	compl(t)=alunos & sala & obs
					celula1=celula1 & "<b>" & rs1("materia") & "</b><br>" & rs1("nome") & " (" & rs1("chapa1") & ")" & "<br>"
					celula2=celula2 & rs1("materia") & """ & chr(10) & """ & rs1("nome") & " (" & rs1("chapa1") & ")" & """ & chr(10) & """
					celula2c=""
					if rs1("chapa2")<>"" then
					    sql="select nome from grades_aux_prof2 where chapa='" & rs1("chapa2") & "' "
					    rs2.Open sql, ,adOpenStatic, adLockReadOnly: professor2=rs2("nome") : rs2.close
						celula1=celula1 & "2: " & professor2 & " (" & rs1("chapa2") & ")<br>"
						celula2c=professor2 & " (" & chapa2 & ")"
					end if
					fundo="#FFFFFF"
				rs1.movenext
				loop
				else 'rs1.recordcount
					celula1="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					celula2="" : celula2b="" : celula2c="" : fundo="#CCCCCC" : compl(t)=""
				end if
				rs1.close
				else
					celula1="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					celula2="" : celula2b="" : celula2c="" : fundo="#CCCCCC" : compl(t)=""
				end if
				disciplina(t)=celula1
				planilha(t)=celula2
				planilha3(t)=celula2c
				backfundo(t)=fundo
			next 'sturma t
%>
	<td class="campoa"r nowrap <%response.write "style='border-top:2px solid #000000'"%> ><%=rs("descricao")%></td>
<%
coluna=2
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & rs("descricao") & """"
			for e=0 to ubound(sturma)
%>					
	<td class="campor" valign=top style="background-color: <%=backfundo(e)%>" <%response.write "style='border-top:2px solid #000000'"%>
	 rowspan="<%=linhas%>"><%=disciplina(e)%><%=compl(e)%>
	</td>
<%
			coluna=coluna+1
			leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & planilha(e) & """"
			if planilha3(e)<>"" then
				leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = objXL.Cells(" & linha & ", " & coluna & ").Value & chr(10) & """ & planilha3(e) & """"
			end if
			next 'e - 
			linha=linha+1

		ultimodia=rs("diasem")
%>
</tr>
<%
		rs.movenext:loop
		rs.close
%>
</table>
<!-- <DIV style="page-break-after:always"></DIV> -->
<%
			if b<ubound(turnograde) then response.write "<DIV style=""page-break-after:always""></DIV>"
response.write "<DIV style=""page-break-after:always""></DIV>"
		next 'turnograde b
		'erase turnograde
		'erase pula
	next 'periodoletivo a
leitura.close
set leitura=nothing
set arquivo=nothing
termino=now()
duracao=(termino-inicio)
'Response.write "<p class=realce><font size=1> Inicio: " & inicio & " Termino: " & termino & " Duracao: " & formatdatetime(duracao,3) & "</font></p>"
%>
<p style="font-size: 8pt;color:gray;margin-top: 0; margin-bottom: 0"><%=right(formatdatetime(duracao,3),5)%>
<a href="../temp/<%=nomefile%>"><img src="../images/msexcel.gif" width="16" height="16" border="0" alt="Clique aqui para exportar este quadro para planilha do Excel"></a></p>
<%
end if 'finaliza 1

set rs=nothing
set rs1=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>