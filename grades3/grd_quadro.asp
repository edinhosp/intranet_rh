<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a90")="N" or session("a90")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Grades - Quadro de Horário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<script language="JavaScript" type="text/javascript" src="../date.js"></script>
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

dim conexao, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsq=server.createobject ("ADODB.Recordset")
Set rsq.ActiveConnection = conexao

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
	<td class=titulo>Escolha o Curso Livre</td></tr>
<tr>
	<td class=titulo><select size="1" name="codcur">
	<option value="0" selected>Selecione um curso</option>
<%
sqla="SELECT gc.coddoc, gc.CURSO " & _
"FROM grades_3 AS gr INNER JOIN g2cursoeve AS gc ON gr.coddoc = gc.coddoc " & _
"GROUP BY gc.coddoc, gc.CURSO ORDER BY gc.CURSO; "

if session("usuariomaster")    then
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
	dim disciplina(10,70), planilha(10,70), planilha2(10,70), planilha3(10,70), backfundo(10,70), compl(10,70)
	for t1=0 to 10
		for t2=0 to 70
			disciplina(t1,t2)=t1*t2+rnd(3)
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

	sql1="SELECT c.coddoc, g.curso, g.perlet, g.perletsg " & _
	"FROM grades_3 g, g2cursoeve c WHERE c.coddoc=g.coddoc " & _
	"AND g.deletada=0 AND '" & datarel & "' Between [inicio] And [termino] " & _
	"GROUP BY c.coddoc, g.curso, g.perlet, g.perletsg HAVING c.coddoc='" & codcur & "' ORDER BY g.perlet;"
	'response.write "<br>Etapa 1: " & sql1               '****************************
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	'response.write "<br>Resul 1: " & rs.recordcount     '****************************
	rs.movefirst
	do while not rs.eof
		redim preserve periodoletivo(numero):redim preserve periodoletivosg(numero)
		periodoletivo(numero)=rs("perlet"):periodoletivosg(numero)=rs("perletsg")
		curso=rs("curso")
	rs.movenext
	numero=numero+1
	loop
	rs.close

	for a=0 to ubound(periodoletivo)
		numero=0
		sql2="SELECT turno, descturno " & _
		"FROM grades_3 g, g2cursoeve c, eturnos e " & _
		"WHERE c.coddoc=g.coddoc AND deletada=0 and g.turno=e.codturno " & _
		"AND c.coddoc='" & codcur & "' AND perlet Like '" & periodoletivo(a) & "' AND '" & datarel & "' between [inicio] and [termino] " & _
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
			sql5="SELECT serie, turma, cast([serie] as varchar)+[turma] AS STURMA " & _
			"FROM grades_3 g, g2cursoeve c " & _
			"WHERE c.coddoc=g.coddoc " & _
			"AND deletada=0 AND c.coddoc='" & codcur & "' AND perlet Like '" & periodoletivo(a) & "%' AND '" & datarel & "' between [inicio] and [termino] " & _
			"and turno=" & turnograde(b) & " " & _
			"GROUP BY serie, turma, cast([serie] as varchar)+[turma] "
'			"AND turno=" & turnograde(b) & " " & _
			'response.write "<br>Etapa 3: " & sql5              '****************************
			rs.Open sql5, ,adOpenStatic, adLockReadOnly
			'response.write "<br>Resul 3: " & rs.recordcount    '****************************
			rs.movefirst
			do while not rs.eof
				redim preserve sturma(numero)
				sturma(numero)=rs("serie") & rs("turma"):'response.write sturma(numero) & " "
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
			numero=0
			sql3="SELECT diasem FROM grades_3 g, g2cursoeve c " & _
			"WHERE c.coddoc=g.coddoc " & _
			"AND deletada=0 AND '" & datarel & "' between [inicio] and [termino] " & _
			"AND c.coddoc='" & codcur & "' AND perlet Like '" & periodoletivo(a) & "' AND turno=" & turnograde(b) & " " & _
			"GROUP BY diasem ORDER BY diasem "
			'response.write "<br>Etapa 4: " & sql3              '****************************
			rs.Open sql3, ,adOpenStatic, adLockReadOnly
			'response.write "<br>Resul 4: " & rs.recordcount    '****************************
			rs.movefirst
			do while not rs.eof
				redim preserve diasemana(numero)
				diasemana(numero)=rs("diasem")
			rs.movenext
			numero=numero+1
			loop
			rs.close
%>
<tr>
<%
			for c=0 to ubound(diasemana)
				numero=0
				'sql4="SELECT horini, horfim FROM grades WHERE deletada=0 AND #" & datarel & "# between [inicio] and [termino] " & _
				'"AND codcur=" & codcur & " AND perlet Like '" & periodoletivo(a) & "' AND turno=" & turnograde(b) & " " & _
				'"AND diasem=" & diasemana(c) & " " & _
				'"GROUP BY horini, horfim ORDER BY horini "
				sql4="select horini, horfim from grd_defhor where codtn=" & turnograde(b) & " and codds=" & diasemana(c) & " " & _
				"group by horini, horfim order by horini "
				'response.write "<br>Etapa 5: " & sql4              '****************************
				rs.Open sql4, ,adOpenStatic, adLockReadOnly
				'response.write "<br>Resul 5: " & rs.recordcount    '****************************
				rs.movefirst
				do while not rs.eof
					redim preserve horaini(numero)
					redim preserve horafim(numero)
					horaini(numero)=rs("horini")
					horafim(numero)=rs("horfim")
				rs.movenext
				numero=numero+1
				loop
				rs.close
				temp=ubound(horaini)
%>
	<td class="campoa"r rowspan="<%=ubound(horaini)+1%>" style="border-top:2px solid #000000" ><%=weekdayname(diasemana(c))%></td>
<%
linhainicio=linha
coluna=1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & weekdayname(diasemana(c)) & """"
'linha=linha+ubound(horaini)

				for d=0 to ubound(horaini)
				'for d=1 to 6
					numero=0
					for e=0 to ubound(sturma)
						celula1="1":celula3="1"
						codcur=codcur
						perlet=periodoletivo(a):perletsg=periodoletivosg(a)
						turno=turnograde(b)
						diasem=diasemana(c)
						horini=horaini(d)
						horfim=horafim(d)
						serie=left(sturma(e),len(sturma(e))-1)
						turma=right(sturma(e),1)
						sql6="select g.coddoc, codmat, materia, chapa1, chapa2, nome, codsala, juntar, jturma, dividir, codtur, p.codcur FROM grades_3 g, " & _
						"g2cursoeve c, grades_per p, grades_gc gc, grades_aux_prof as f WHERE deletada=0 AND chapa1=chapa " & _
						"AND g.coddoc=c.coddoc and p.coddoc=g.coddoc and p.perlet=g.perlet and gc.serie=g.serie and gc.coddoc=p.coddoc and gc.gc=p.gc and gc.perlet=p.perlet " & _
						" AND '" & datarel & "' between [inicio] and [termino] " & _
						" AND c.coddoc='" & codcur & "' AND g.perlet='" & perlet & "'" & " AND turno=" & turno & _
						" AND g.serie=" & serie & "" & " AND turma='" & turma & "'" & " AND diasem=" & diasem & _
						" AND a" & d+1 & "=1 and ativo=1 "
						'response.write "<br>Etapa 6: " & sql6 & "<br>"
						rs.Open sql6, ,adOpenStatic, adLockReadOnly
						if rs.recordcount>0 then
						celula1="":celula2="":celula2b="":celula2c=""						
						rs.movefirst
						do while not rs.eof
							codmat=rs("codmat")
							materia=rs("materia")
							chapa1=rs("chapa1")
							nome=rs("nome")
							chapa2=rs("chapa2")
							codtur=rs("codtur")
							codcurrm=rs("codcur")
							'*******alunos******
	'if codcur=260 then
	'	if perlet="2004/0" and turma="B" then perlet="2003/2"
	'end if
	'sql3="SELECT uma.STATUS, USITMAT_1.DESCRICAO, Count(umc.MATALUNO) AS alunos " & _
	'antes do having "GROUP BY uma.STATUS, USITMAT_1.DESCRICAO " & _
	if request.form("naluno")="ON" then
	sql3="SELECT Count(umc.MATALUNO) AS alunos " & _
	"FROM corporerm.dbo.USITMAT AS usm1 INNER JOIN ((((corporerm.dbo.UMATRICPL umc INNER JOIN corporerm.dbo.UMATALUN uma ON (umc.GRADE=uma.GRADE) AND (umc.CODPER=uma.CODPER) AND (umc.CODCUR=uma.CODCUR) AND (umc.PERLETIVO=uma.PERLETIVO) AND (umc.MATALUNO=uma.MATALUNO) AND (umc.CODCOLIGADA=uma.CODCOLIGADA) AND (umc.CODFILIAL=uma.CODFILIAL)) " & _
	"INNER JOIN corporerm.dbo.UMATERIAS um ON (uma.CODCOLIGADA=um.CODCOLIGADA) AND (uma.CODMAT=um.CODMAT)) " & _
	"INNER JOIN corporerm.dbo.UALUCURSO a ON (umc.GRADE=a.GRADE) AND (umc.CODPER=a.CODPER) AND (umc.CODCUR=a.CODCUR) AND (umc.MATALUNO=a.MATALUNO) AND (umc.CODCOLIGADA=a.CODCOLIGADA)) " & _
	"INNER JOIN corporerm.dbo.USITMAT usm ON a.STATUS=usm.CODSITMAT) ON usm1.CODSITMAT=uma.STATUS " & _
	"WHERE umc.PERLETIVO='" & perletsg & "' " & _
	"AND umc.CODCUR=" & codcurrm & " " & _
	"AND CODTUR='" & codtur & "' " & _
	"AND uma.CODMAT='" & codmat & "' " & _
	"and uma.STATUS In ('01','07','08','09','10','18','19','20','47','48','46','70') "
	'"HAVING uma.STATUS In ('01') "
	'"AND uma.CODMAT='" & codmat & "' AND UALUCURSO.STATUS='01' " & _
	'response.write sql3
	rsq.Open sql3, ,adOpenStatic, adLockReadOnly
	if rsq.recordcount>0 then alunos="Alunos: " & rsq("alunos") & "<br>" else alunos="erro " & codtur
	rsq.close
	else
		alunos="Alunos: --<br>"
	end if
	if isnull(rs("codsala")) then sala="" else sala="Sala: " & rs("codsala") & "<br>"
	if rs("juntar")=-1 then
		obs="<font color=blue>Junta turma " & rs("jturma") & "<br>"
	elseif rs("dividir")=-1 then
		obs="<font color=red>Divide turma" & "<br>"
	else
		obs=""
	end if
	compl(d,e)=alunos & sala & obs
							celula1=celula1 & "<b>" & materia & "</b><br>" & nome & " (" & chapa1 & ")" & "<br>"
'if rs.recordcount>1 then celula1=celula1 & compl(d,e)
							celula2=celula2 & materia & """ & chr(10) & """ & nome & " (" & chapa1 & ")" & """ & chr(10) & """
							'celula2b=nome & " (" & chapa1 & ")"
							celula2c=""
							if rs("chapa2")<>"" then
							    sql="select nome from grades_aux_prof where chapa='" & rs("chapa2") & "' "
							    rs2.Open sql, ,adOpenStatic, adLockReadOnly
							    professor2=rs2("nome")
							    rs2.close
								celula1=celula1 & "2: " & professor2 & " (" & chapa2 & ")<br>"
								celula2c=professor2 & " (" & chapa2 & ")"
							end if
							fundo="#FFFFFF"
						rs.movenext
						loop
						else
							celula1="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
							celula2=""
							celula2b=""
							celula2c=""
							fundo="#CCCCCC"
							compl(d,e)=""
						end if
						rs.close
						disciplina(d,e)=celula1
						planilha(d,e)=celula2
						'planilha2(d,e)=celula2b
						planilha3(d,e)=celula2c
						backfundo(d,e)=fundo
					next 'sturma e

				next 'horainicio d
					
				for d=0 to ubound(horaini)
%>
	<td class="campoa"r nowrap <%if d=0 then response.write "style='border-top:2px solid #000000'"%> ><%=formatdatetime(horaini(d),4)%>-<%=formatdatetime(horafim(d),4)%></td>
<%
coluna=2
'linha=linhainicio
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & formatdatetime(horaini(d),4) & "-" & formatdatetime(horafim(d),4) & """"
					'for e=0 to ubound(sturma)
					'	if e<>ubound(sturma) then
					'	'if coluna(e)="" then coluna(e)=0
					'	for d2=d+1 to ubound(horaini)
					'		if disciplina(d2,e)=disciplina(d,e) then coluna(e)=coluna(e)+1
					'	next
					'	end if
					'next 
					for e=0 to ubound(sturma)
						select case d
							case 0
								linhas=1
								if disciplina(0,e)=disciplina(1,e) then linhas=2
								if disciplina(0,e)=disciplina(2,e) and disciplina(1,e)=disciplina(2,e) then linhas=3
								if disciplina(0,e)=disciplina(3,e) and disciplina(1,e)=disciplina(3,e) and disciplina(2,e)=disciplina(3,e) then linhas=4
								if disciplina(0,e)=disciplina(4,e) and disciplina(1,e)=disciplina(4,e) and disciplina(2,e)=disciplina(4,e) and disciplina(3,e)=disciplina(4,e) then linhas=5
								if disciplina(0,e)=disciplina(5,e) and disciplina(1,e)=disciplina(5,e) and disciplina(2,e)=disciplina(5,e) and disciplina(3,e)=disciplina(5,e) and disciplina(4,e)=disciplina(5,e) then linhas=6
							case 1
								linhas=1
								if disciplina(1,e)=disciplina(2,e) then linhas=2
								if disciplina(1,e)=disciplina(3,e) and disciplina(2,e)=disciplina(3,e) then linhas=3
								if disciplina(1,e)=disciplina(4,e) and disciplina(2,e)=disciplina(4,e) and disciplina(3,e)=disciplina(4,e) then linhas=4
								if disciplina(1,e)=disciplina(5,e) and disciplina(2,e)=disciplina(5,e) and disciplina(3,e)=disciplina(5,e) and disciplina(4,e)=disciplina(5,e) then linhas=5
								if disciplina(1,e)=disciplina(0,e) then linhas=0
							case 2
								linhas=1
								if disciplina(2,e)=disciplina(3,e) then linhas=2
								if disciplina(2,e)=disciplina(4,e) and disciplina(3,e)=disciplina(4,e) then linhas=3
								if disciplina(2,e)=disciplina(5,e) and disciplina(3,e)=disciplina(5,e) and disciplina(4,e)=disciplina(5,e) then linhas=4
								if disciplina(2,e)=disciplina(1,e) then linhas=0
								'if disciplina(2,e)=disciplina(0,e) then linhas=0
							case 3
								linhas=1
								if disciplina(3,e)=disciplina(4,e) then linhas=2
								if disciplina(3,e)=disciplina(5,e) and disciplina(4,e)=disciplina(5,e) then linhas=3
								if disciplina(3,e)=disciplina(2,e) then linhas=0
								'if disciplina(3,e)=disciplina(1,e) then linhas=0
								'if disciplina(3,e)=disciplina(0,e) then linhas=0
							case 4
								linhas=1
								if disciplina(4,e)=disciplina(5,e) then linhas=2
								if disciplina(4,e)=disciplina(3,e) then linhas=0
								'if disciplina(4,e)=disciplina(2,e) then linhas=0
								'if disciplina(4,e)=disciplina(1,e) then linhas=0
								'if disciplina(4,e)=disciplina(0,e) then linhas=0
							case 5
								linhas=1
								if disciplina(5,e)=disciplina(4,e) then linhas=0
								'if disciplina(5,e)=disciplina(3,e) then linhas=0
								'if disciplina(5,e)=disciplina(2,e) then linhas=0
								'if disciplina(5,e)=disciplina(1,e) then linhas=0
								'if disciplina(5,e)=disciplina(0,e) then linhas=0
						end select
					if linhas>0 then
%>					
	<td class="campor" valign=top style="background-color: <%=backfundo(d,e)%>" <%if d=0 then response.write "style='border-top:2px solid #000000'"%>
	 rowspan="<%=linhas%>"><%=disciplina(d,e)%><%=compl(d,e)%>
	</td>
<%
					end if
coluna=coluna+1
'leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & planilha(d,e) & """ & chr(10) & """ & planilha2(d,e) & """"
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & planilha(d,e) & """"
if planilha3(d,e)<>"" then
	leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = objXL.Cells(" & linha & ", " & coluna & ").Value & chr(10) & """ & planilha3(d,e) & """"
end if
					next 'sturma e
linha=linha+1
%>
</tr>
<%
				next 'horainicio d
				'erase pula

			next 'diasemana c
	sql0="SELECT c.CURSO, diretor, coordenador, adjunto, chefedepto, secretaria FROM grades_per g, " & _
	"g2cursoeve c WHERE g.coddoc=c.coddoc " & _
	"AND c.coddoc='" & codcur & "' and perlet='" & periodoletivo(a) & "' "
	rs.Open sql0, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
		diretor=rs("diretor")
		coordenador=rs("coordenador")
		chefedepto=rs("chefedepto")
		secretaria=rs("secretaria")
		if codcur<>"DIR" and (turnograde(b)="1" or turnograde(b)="2") and left(periodoletivo(a),4)>="2004" and isnull(diretor) then
			diretor=" "
		end if
		if codcur<>"DIR" and (turnograde(b)="3") and left(periodoletivo(a),4)>="2004" and isnull(diretor) then
			diretor=" "
		end if
	else
		diretor="-"
		coordenador="-"
		chefedepto="-"
		secretaria="-"
	end if
	rs.close	
			
%>
<tr>
	<td class="campor" colspan="<%=3+ubound(sturma)+1%>" style="border-top:2px solid #000000" >
	Diretor: <%=diretor%><br>Coordenador do curso: <%=coordenador%>
	<br>Chefe de Departamento: <%=chefedepto%>
	</td>
</tr>
</table>
<!-- <DIV style="page-break-after:always"></DIV> -->
<%
coluna=1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = ""Diretor: " & diretor & """"
linha=linha+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = ""Coordenador do curso: " & coordenador & """"
linha=linha+1
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = ""Chefe de Departamento: " & chefedepto & """"
linha=linha+1

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
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>