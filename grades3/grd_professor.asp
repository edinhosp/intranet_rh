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
<title>Grades - Quadro de Horário do Professor</title>
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
sqla="SELECT F.CHAPA, F.NOME FROM grades_3ch g INNER JOIN " & _
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
		</select>
	</td>
</tr>
<tr>
	<td class=titulo>Horário vigente na data de &nbsp;<input type=text name="database" value=<%=database%> size=10 class=a onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')">
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
	dim disciplina(15,70), planilha(15,70), planilha2(15,70), planilha3(15,70), backfundo(15,70), compl(15,70)
	for t1=0 to 15
		for t2=0 to 70
			disciplina(t1,t2)=t1*t2+rnd(3)
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

	sql0="select nome from grades_aux_prof where chapa='" & chapa & "' "
	rs.Open sql0, ,adOpenStatic, adLockReadOnly
	nomechapa=rs("nome")
	rs.close

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
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & chapa & " - " & nomechapa & """"

linha=linha+1

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
			'for c=0 to ubound(diasemana)
				numero=0
				sql4="SELECT descricao FROM grades_3ch WHERE deletada=0 " & _
				"AND '" & datarel & "' between inicio and termino " & _
				"AND chapa1='" & chapa & "' " & _
				"GROUP BY descricao ORDER BY descricao "
				'response.write "<br>" & sql4                        '*************************************
				rs.Open sql4, ,adOpenStatic, adLockReadOnly
				'response.write "<br>" & rs.recordcount              '*************************************
				rs.movefirst
				do while not rs.eof
					redim preserve horaini(numero)
					'redim preserve horafim(numero)
					horaini(numero)=rs("descricao")
					'horafim(numero)=rs("horfim")
				rs.movenext
				numero=numero+1
				loop
				rs.close
				temp=ubound(horaini)

linhainicio=linha
coluna=1
'linha=linha+ubound(horaini)
				for d=0 to ubound(horaini)
					numero=0
					
					for e=0 to ubound(diasemana)
						celula1="1":celula3="1"
						chapa=chapa
						diasem=diasemana(e)
						horini=horaini(d)
						'horfim=horafim(d)
						sql6="SELECT distinct g.perlet, g.perletsg, p.codcur, g.coddoc, g.curso, codtur, codmat, materia, g.serie, turma, chapa1, nome, codsala, juntar, jturma, dividir FROM grades_3ch g, " & _
						"grades_per p, grades_gc gc, " & _
						"grades_aux_prof as f WHERE deletada=0 AND g.chapa1=f.chapa " & _
						"and p.coddoc=gc.coddoc and p.gc=gc.gc and p.perlet=gc.perlet and g.coddoc=p.coddoc and g.perlet=p.perlet and g.serie=gc.serie " & _
						" AND '" & datarel & "' between inicio and termino " & _
						" AND diasem=" & diasem & _
						" AND descricao='" & horini & "' and chapa='" & chapa & "' "
						'response.write "<br><br>" &sql6
						rs.Open sql6, ,adOpenStatic, adLockReadOnly
						celula1=""
						if rs.recordcount>0 then
							do while not rs.eof
							codcur=rs("coddoc")
							codcurrm=rs("codcur")
							curso=rs("curso")
							codmat=rs("codmat")
							materia=rs("materia")
							serie=rs("serie")
							turma=rs("turma")
							chapa1=rs("chapa1")
							nome=rs("nome")
							perlet=rs("perlet"):perletsg=rs("perletsg")
							codtur=rs("codtur")
							sala=rs("codsala"):if sala="0" then sala=""
'	if codcur=260 then
'		if perlet="2004/0" and turma="B" then perlet="2003/2"
'	end if
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
	rsq.Open sql3, ,adOpenStatic, adLockReadOnly
	if rsq.recordcount>0 then alunos="Alunos: " & rsq("alunos") & "<br>" else alunos=""							
	rsq.close
	else
		alunos="Alunos: --<br>"
	end if
	if rs("juntar")=-1 then
		obs="<font color=blue>Junta turma " & rs("jturma") & "<br>"
	elseif rs("dividir")=-1 then
		obs="<font color=red>Divide turma" & "<br>"
	else
		obs=""
	end if
	compl(d,e)=alunos & obs
							celula1=celula1 & "<b>" & curso & " / " & serie & turma & "</b><br>" & materia & "<br>Sala " & sala
							celula2=curso & " / " & serie & turma 
							celula2b=materia & duplicado
							celula2c=sala
							rs.movenext
							celula1=celula1 & "<br>"
							loop
							fundo="#FFFFFF"
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
						planilha2(d,e)=celula2b
						planilha3(d,e)=celula2c
						backfundo(d,e)=fundo
					next 'sturma e

				next 'horainicio d
					
				for d=0 to ubound(horaini)
%>
	<td class="campoa"r nowrap><%=horaini(d)%></td>
<%
coluna=1
'linha=linhainicio
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & horaini(d) & """"
					'for e=0 to ubound(sturma)
					'	if e<>ubound(sturma) then
					'	'if coluna(e)="" then coluna(e)=0
					'	for d2=d+1 to ubound(horaini)
					'		if disciplina(d2,e)=disciplina(d,e) then coluna(e)=coluna(e)+1
					'	next
					'	end if
					'next 
					for e=0 to ubound(diasemana)
						select case d
							case 0
								linhas=1
								if disciplina(0,e)=disciplina(1,e) then linhas=2
								if disciplina(0,e)=disciplina(2,e) and disciplina(1,e)=disciplina(2,e) then linhas=3
								if disciplina(0,e)=disciplina(3,e) and disciplina(1,e)=disciplina(3,e) and disciplina(2,e)=disciplina(3,e) then linhas=4
								if disciplina(0,e)=disciplina(4,e) and disciplina(1,e)=disciplina(4,e) and disciplina(2,e)=disciplina(4,e) and disciplina(3,e)=disciplina(4,e) then linhas=5
								if disciplina(0,e)=disciplina(5,e) and disciplina(1,e)=disciplina(5,e) and disciplina(2,e)=disciplina(5,e) and disciplina(3,e)=disciplina(5,e) and disciplina(4,e)=disciplina(5,e) then linhas=6
								if disciplina(0,e)=disciplina(6,e) and disciplina(1,e)=disciplina(6,e) and disciplina(2,e)=disciplina(6,e) and disciplina(3,e)=disciplina(6,e) and disciplina(4,e)=disciplina(6,e) and disciplina(5,e)=disciplina(6,e) then linhas=7
								if disciplina(0,e)=disciplina(7,e) and disciplina(1,e)=disciplina(7,e) and disciplina(2,e)=disciplina(7,e) and disciplina(3,e)=disciplina(7,e) and disciplina(4,e)=disciplina(7,e) and disciplina(5,e)=disciplina(7,e) and disciplina(6,e)=disciplina(7,e) then linhas=8
								if disciplina(0,e)=disciplina(8,e) and disciplina(1,e)=disciplina(8,e) and disciplina(2,e)=disciplina(8,e) and disciplina(3,e)=disciplina(8,e) and disciplina(4,e)=disciplina(8,e) and disciplina(5,e)=disciplina(8,e) and disciplina(6,e)=disciplina(8,e) and disciplina(7,e)=disciplina(8,e) then linhas=9
								if disciplina(0,e)=disciplina(9,e) and disciplina(1,e)=disciplina(9,e) and disciplina(2,e)=disciplina(9,e) and disciplina(3,e)=disciplina(9,e) and disciplina(4,e)=disciplina(9,e) and disciplina(5,e)=disciplina(9,e) and disciplina(6,e)=disciplina(9,e) and disciplina(7,e)=disciplina(9,e) and disciplina(8,e)=disciplina(9,e) then linhas=10
								if disciplina(0,e)=disciplina(10,e) and disciplina(1,e)=disciplina(10,e) and disciplina(2,e)=disciplina(10,e) and disciplina(3,e)=disciplina(10,e) and disciplina(4,e)=disciplina(10,e) and disciplina(5,e)=disciplina(10,e) and disciplina(6,e)=disciplina(10,e) and disciplina(7,e)=disciplina(10,e) and disciplina(8,e)=disciplina(10,e) and disciplina(9,e)=disciplina(10,e) then linhas=11
								if disciplina(0,e)=disciplina(11,e) and disciplina(1,e)=disciplina(11,e) and disciplina(2,e)=disciplina(11,e) and disciplina(3,e)=disciplina(11,e) and disciplina(4,e)=disciplina(11,e) and disciplina(5,e)=disciplina(11,e) and disciplina(6,e)=disciplina(11,e) and disciplina(7,e)=disciplina(11,e) and disciplina(8,e)=disciplina(11,e) and disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=12
								if disciplina(0,e)=disciplina(12,e) and disciplina(1,e)=disciplina(12,e) and disciplina(2,e)=disciplina(12,e) and disciplina(3,e)=disciplina(12,e) and disciplina(4,e)=disciplina(12,e) and disciplina(5,e)=disciplina(12,e) and disciplina(6,e)=disciplina(12,e) and disciplina(7,e)=disciplina(12,e) and disciplina(8,e)=disciplina(12,e) and disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=13
								if disciplina(0,e)=disciplina(13,e) and disciplina(1,e)=disciplina(13,e) and disciplina(2,e)=disciplina(13,e) and disciplina(3,e)=disciplina(13,e) and disciplina(4,e)=disciplina(13,e) and disciplina(5,e)=disciplina(13,e) and disciplina(6,e)=disciplina(13,e) and disciplina(7,e)=disciplina(13,e) and disciplina(8,e)=disciplina(13,e) and disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=14
								if disciplina(0,e)=disciplina(14,e) and disciplina(1,e)=disciplina(14,e) and disciplina(2,e)=disciplina(14,e) and disciplina(3,e)=disciplina(14,e) and disciplina(4,e)=disciplina(14,e) and disciplina(5,e)=disciplina(14,e) and disciplina(6,e)=disciplina(14,e) and disciplina(7,e)=disciplina(14,e) and disciplina(8,e)=disciplina(14,e) and disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=15
								if disciplina(0,e)=disciplina(15,e) and disciplina(1,e)=disciplina(15,e) and disciplina(2,e)=disciplina(15,e) and disciplina(3,e)=disciplina(15,e) and disciplina(4,e)=disciplina(15,e) and disciplina(5,e)=disciplina(15,e) and disciplina(6,e)=disciplina(15,e) and disciplina(7,e)=disciplina(15,e) and disciplina(8,e)=disciplina(15,e) and disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=16
							case 1
								linhas=1
								if disciplina(1,e)=disciplina(2,e) then linhas=2
								if disciplina(1,e)=disciplina(3,e) and disciplina(2,e)=disciplina(3,e) then linhas=3
								if disciplina(1,e)=disciplina(4,e) and disciplina(2,e)=disciplina(4,e) and disciplina(3,e)=disciplina(4,e) then linhas=4
								if disciplina(1,e)=disciplina(5,e) and disciplina(2,e)=disciplina(5,e) and disciplina(3,e)=disciplina(5,e) and disciplina(4,e)=disciplina(5,e) then linhas=5
								if disciplina(1,e)=disciplina(6,e) and disciplina(2,e)=disciplina(6,e) and disciplina(3,e)=disciplina(6,e) and disciplina(4,e)=disciplina(6,e) and disciplina(5,e)=disciplina(6,e) then linhas=6
								if disciplina(1,e)=disciplina(7,e) and disciplina(2,e)=disciplina(7,e) and disciplina(3,e)=disciplina(7,e) and disciplina(4,e)=disciplina(7,e) and disciplina(5,e)=disciplina(7,e) and disciplina(6,e)=disciplina(7,e) then linhas=7
								if disciplina(1,e)=disciplina(8,e) and disciplina(2,e)=disciplina(8,e) and disciplina(3,e)=disciplina(8,e) and disciplina(4,e)=disciplina(8,e) and disciplina(5,e)=disciplina(8,e) and disciplina(6,e)=disciplina(8,e) and disciplina(7,e)=disciplina(8,e) then linhas=8
								if disciplina(1,e)=disciplina(9,e) and disciplina(2,e)=disciplina(9,e) and disciplina(3,e)=disciplina(9,e) and disciplina(4,e)=disciplina(9,e) and disciplina(5,e)=disciplina(9,e) and disciplina(6,e)=disciplina(9,e) and disciplina(7,e)=disciplina(9,e) and disciplina(8,e)=disciplina(9,e) then linhas=9
								if disciplina(1,e)=disciplina(10,e) and disciplina(2,e)=disciplina(10,e) and disciplina(3,e)=disciplina(10,e) and disciplina(4,e)=disciplina(10,e) and disciplina(5,e)=disciplina(10,e) and disciplina(6,e)=disciplina(10,e) and disciplina(7,e)=disciplina(10,e) and disciplina(8,e)=disciplina(10,e) and disciplina(9,e)=disciplina(10,e) then linhas=10
								if disciplina(1,e)=disciplina(11,e) and disciplina(2,e)=disciplina(11,e) and disciplina(3,e)=disciplina(11,e) and disciplina(4,e)=disciplina(11,e) and disciplina(5,e)=disciplina(11,e) and disciplina(6,e)=disciplina(11,e) and disciplina(7,e)=disciplina(11,e) and disciplina(8,e)=disciplina(11,e) and disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=11
								if disciplina(1,e)=disciplina(12,e) and disciplina(2,e)=disciplina(12,e) and disciplina(3,e)=disciplina(12,e) and disciplina(4,e)=disciplina(12,e) and disciplina(5,e)=disciplina(12,e) and disciplina(6,e)=disciplina(12,e) and disciplina(7,e)=disciplina(12,e) and disciplina(8,e)=disciplina(12,e) and disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=12
								if disciplina(1,e)=disciplina(13,e) and disciplina(2,e)=disciplina(13,e) and disciplina(3,e)=disciplina(13,e) and disciplina(4,e)=disciplina(13,e) and disciplina(5,e)=disciplina(13,e) and disciplina(6,e)=disciplina(13,e) and disciplina(7,e)=disciplina(13,e) and disciplina(8,e)=disciplina(13,e) and disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=13
								if disciplina(1,e)=disciplina(14,e) and disciplina(2,e)=disciplina(14,e) and disciplina(3,e)=disciplina(14,e) and disciplina(4,e)=disciplina(14,e) and disciplina(5,e)=disciplina(14,e) and disciplina(6,e)=disciplina(14,e) and disciplina(7,e)=disciplina(14,e) and disciplina(8,e)=disciplina(14,e) and disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=14
								if disciplina(1,e)=disciplina(15,e) and disciplina(2,e)=disciplina(15,e) and disciplina(3,e)=disciplina(15,e) and disciplina(4,e)=disciplina(15,e) and disciplina(5,e)=disciplina(15,e) and disciplina(6,e)=disciplina(15,e) and disciplina(7,e)=disciplina(15,e) and disciplina(8,e)=disciplina(15,e) and disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=15
								if disciplina(1,e)=disciplina(0,e) then linhas=0
							case 2
								linhas=1
								if disciplina(2,e)=disciplina(3,e) then linhas=2
								if disciplina(2,e)=disciplina(4,e) and disciplina(3,e)=disciplina(4,e) then linhas=3
								if disciplina(2,e)=disciplina(5,e) and disciplina(3,e)=disciplina(5,e) and disciplina(4,e)=disciplina(5,e) then linhas=4
								if disciplina(2,e)=disciplina(6,e) and disciplina(3,e)=disciplina(6,e) and disciplina(4,e)=disciplina(6,e) and disciplina(5,e)=disciplina(6,e) then linhas=5
								if disciplina(2,e)=disciplina(7,e) and disciplina(3,e)=disciplina(7,e) and disciplina(4,e)=disciplina(7,e) and disciplina(5,e)=disciplina(7,e) and disciplina(6,e)=disciplina(7,e) then linhas=6
								if disciplina(2,e)=disciplina(8,e) and disciplina(3,e)=disciplina(8,e) and disciplina(4,e)=disciplina(8,e) and disciplina(5,e)=disciplina(8,e) and disciplina(6,e)=disciplina(8,e) and disciplina(7,e)=disciplina(8,e) then linhas=7
								if disciplina(2,e)=disciplina(9,e) and disciplina(3,e)=disciplina(9,e) and disciplina(4,e)=disciplina(9,e) and disciplina(5,e)=disciplina(9,e) and disciplina(6,e)=disciplina(9,e) and disciplina(7,e)=disciplina(9,e) and disciplina(8,e)=disciplina(9,e) then linhas=8
								if disciplina(2,e)=disciplina(10,e) and disciplina(3,e)=disciplina(10,e) and disciplina(4,e)=disciplina(10,e) and disciplina(5,e)=disciplina(10,e) and disciplina(6,e)=disciplina(10,e) and disciplina(7,e)=disciplina(10,e) and disciplina(8,e)=disciplina(10,e) and disciplina(9,e)=disciplina(10,e) then linhas=9
								if disciplina(2,e)=disciplina(11,e) and disciplina(3,e)=disciplina(11,e) and disciplina(4,e)=disciplina(11,e) and disciplina(5,e)=disciplina(11,e) and disciplina(6,e)=disciplina(11,e) and disciplina(7,e)=disciplina(11,e) and disciplina(8,e)=disciplina(11,e) and disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=10
								if disciplina(2,e)=disciplina(12,e) and disciplina(3,e)=disciplina(12,e) and disciplina(4,e)=disciplina(12,e) and disciplina(5,e)=disciplina(12,e) and disciplina(6,e)=disciplina(12,e) and disciplina(7,e)=disciplina(12,e) and disciplina(8,e)=disciplina(12,e) and disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=11
								if disciplina(2,e)=disciplina(13,e) and disciplina(3,e)=disciplina(13,e) and disciplina(4,e)=disciplina(13,e) and disciplina(5,e)=disciplina(13,e) and disciplina(6,e)=disciplina(13,e) and disciplina(7,e)=disciplina(13,e) and disciplina(8,e)=disciplina(13,e) and disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=12
								if disciplina(2,e)=disciplina(14,e) and disciplina(3,e)=disciplina(14,e) and disciplina(4,e)=disciplina(14,e) and disciplina(5,e)=disciplina(14,e) and disciplina(6,e)=disciplina(14,e) and disciplina(7,e)=disciplina(14,e) and disciplina(8,e)=disciplina(14,e) and disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=13
								if disciplina(2,e)=disciplina(15,e) and disciplina(3,e)=disciplina(15,e) and disciplina(4,e)=disciplina(15,e) and disciplina(5,e)=disciplina(15,e) and disciplina(6,e)=disciplina(15,e) and disciplina(7,e)=disciplina(15,e) and disciplina(8,e)=disciplina(15,e) and disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=14
								if disciplina(2,e)=disciplina(1,e) then linhas=0
								'if disciplina(2,e)=disciplina(0,e) then linhas=0
							case 3
								linhas=1
								if disciplina(3,e)=disciplina(4,e) then linhas=2
								if disciplina(3,e)=disciplina(5,e) and disciplina(4,e)=disciplina(5,e) then linhas=3
								if disciplina(3,e)=disciplina(6,e) and disciplina(4,e)=disciplina(6,e) and disciplina(5,e)=disciplina(6,e) then linhas=4
								if disciplina(3,e)=disciplina(7,e) and disciplina(4,e)=disciplina(7,e) and disciplina(5,e)=disciplina(7,e) and disciplina(6,e)=disciplina(7,e) then linhas=5
								if disciplina(3,e)=disciplina(8,e) and disciplina(4,e)=disciplina(8,e) and disciplina(5,e)=disciplina(8,e) and disciplina(6,e)=disciplina(8,e) and disciplina(7,e)=disciplina(8,e) then linhas=6
								if disciplina(3,e)=disciplina(9,e) and disciplina(4,e)=disciplina(9,e) and disciplina(5,e)=disciplina(9,e) and disciplina(6,e)=disciplina(9,e) and disciplina(7,e)=disciplina(9,e) and disciplina(8,e)=disciplina(9,e) then linhas=7
								if disciplina(3,e)=disciplina(10,e) and disciplina(4,e)=disciplina(10,e) and disciplina(5,e)=disciplina(10,e) and disciplina(6,e)=disciplina(10,e) and disciplina(7,e)=disciplina(10,e) and disciplina(8,e)=disciplina(10,e) and disciplina(9,e)=disciplina(10,e) then linhas=8
								if disciplina(3,e)=disciplina(11,e) and disciplina(4,e)=disciplina(11,e) and disciplina(5,e)=disciplina(11,e) and disciplina(6,e)=disciplina(11,e) and disciplina(7,e)=disciplina(11,e) and disciplina(8,e)=disciplina(11,e) and disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=9
								if disciplina(3,e)=disciplina(12,e) and disciplina(4,e)=disciplina(12,e) and disciplina(5,e)=disciplina(12,e) and disciplina(6,e)=disciplina(12,e) and disciplina(7,e)=disciplina(12,e) and disciplina(8,e)=disciplina(12,e) and disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=10
								if disciplina(3,e)=disciplina(13,e) and disciplina(4,e)=disciplina(13,e) and disciplina(5,e)=disciplina(13,e) and disciplina(6,e)=disciplina(13,e) and disciplina(7,e)=disciplina(13,e) and disciplina(8,e)=disciplina(13,e) and disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=11
								if disciplina(3,e)=disciplina(14,e) and disciplina(4,e)=disciplina(14,e) and disciplina(5,e)=disciplina(14,e) and disciplina(6,e)=disciplina(14,e) and disciplina(7,e)=disciplina(14,e) and disciplina(8,e)=disciplina(14,e) and disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=12
								if disciplina(3,e)=disciplina(15,e) and disciplina(4,e)=disciplina(15,e) and disciplina(5,e)=disciplina(15,e) and disciplina(6,e)=disciplina(15,e) and disciplina(7,e)=disciplina(15,e) and disciplina(8,e)=disciplina(15,e) and disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=13
								if disciplina(3,e)=disciplina(2,e) then linhas=0
								'if disciplina(3,e)=disciplina(1,e) then linhas=0
								'if disciplina(3,e)=disciplina(0,e) then linhas=0
							case 4
								linhas=1
								if disciplina(4,e)=disciplina(5,e) then linhas=2
								if disciplina(4,e)=disciplina(6,e) and disciplina(5,e)=disciplina(6,e) then linhas=3
								if disciplina(4,e)=disciplina(7,e) and disciplina(5,e)=disciplina(7,e) and disciplina(6,e)=disciplina(7,e) then linhas=4
								if disciplina(4,e)=disciplina(8,e) and disciplina(5,e)=disciplina(8,e) and disciplina(6,e)=disciplina(8,e) and disciplina(7,e)=disciplina(8,e) then linhas=5
								if disciplina(4,e)=disciplina(9,e) and disciplina(5,e)=disciplina(9,e) and disciplina(6,e)=disciplina(9,e) and disciplina(7,e)=disciplina(9,e) and disciplina(8,e)=disciplina(9,e) then linhas=6
								if disciplina(4,e)=disciplina(10,e) and disciplina(5,e)=disciplina(10,e) and disciplina(6,e)=disciplina(10,e) and disciplina(7,e)=disciplina(10,e) and disciplina(8,e)=disciplina(10,e) and disciplina(9,e)=disciplina(10,e) then linhas=7
								if disciplina(4,e)=disciplina(11,e) and disciplina(5,e)=disciplina(11,e) and disciplina(6,e)=disciplina(11,e) and disciplina(7,e)=disciplina(11,e) and disciplina(8,e)=disciplina(11,e) and disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=8
								if disciplina(4,e)=disciplina(12,e) and disciplina(5,e)=disciplina(12,e) and disciplina(6,e)=disciplina(12,e) and disciplina(7,e)=disciplina(12,e) and disciplina(8,e)=disciplina(12,e) and disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=9
								if disciplina(4,e)=disciplina(13,e) and disciplina(5,e)=disciplina(13,e) and disciplina(6,e)=disciplina(13,e) and disciplina(7,e)=disciplina(13,e) and disciplina(8,e)=disciplina(13,e) and disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=10
								if disciplina(4,e)=disciplina(14,e) and disciplina(5,e)=disciplina(14,e) and disciplina(6,e)=disciplina(14,e) and disciplina(7,e)=disciplina(14,e) and disciplina(8,e)=disciplina(14,e) and disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=11
								if disciplina(4,e)=disciplina(15,e) and disciplina(5,e)=disciplina(15,e) and disciplina(6,e)=disciplina(15,e) and disciplina(7,e)=disciplina(15,e) and disciplina(8,e)=disciplina(15,e) and disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=12
								if disciplina(4,e)=disciplina(3,e) then linhas=0
								'if disciplina(4,e)=disciplina(2,e) then linhas=0
								'if disciplina(4,e)=disciplina(1,e) then linhas=0
								'if disciplina(4,e)=disciplina(0,e) then linhas=0
							case 5
								linhas=1
								if disciplina(5,e)=disciplina(6,e) then linhas=2
								if disciplina(5,e)=disciplina(7,e) and disciplina(6,e)=disciplina(7,e) then linhas=3
								if disciplina(5,e)=disciplina(8,e) and disciplina(6,e)=disciplina(8,e) and disciplina(7,e)=disciplina(8,e) then linhas=4
								if disciplina(5,e)=disciplina(9,e) and disciplina(6,e)=disciplina(9,e) and disciplina(7,e)=disciplina(9,e) and disciplina(8,e)=disciplina(9,e) then linhas=5
								if disciplina(5,e)=disciplina(10,e) and disciplina(6,e)=disciplina(10,e) and disciplina(7,e)=disciplina(10,e) and disciplina(8,e)=disciplina(10,e) and disciplina(9,e)=disciplina(10,e) then linhas=6
								if disciplina(5,e)=disciplina(11,e) and disciplina(6,e)=disciplina(11,e) and disciplina(7,e)=disciplina(11,e) and disciplina(8,e)=disciplina(11,e) and disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=7
								if disciplina(5,e)=disciplina(12,e) and disciplina(6,e)=disciplina(12,e) and disciplina(7,e)=disciplina(12,e) and disciplina(8,e)=disciplina(12,e) and disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=8
								if disciplina(5,e)=disciplina(13,e) and disciplina(6,e)=disciplina(13,e) and disciplina(7,e)=disciplina(13,e) and disciplina(8,e)=disciplina(13,e) and disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=9
								if disciplina(5,e)=disciplina(14,e) and disciplina(6,e)=disciplina(14,e) and disciplina(7,e)=disciplina(14,e) and disciplina(8,e)=disciplina(14,e) and disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=10
								if disciplina(5,e)=disciplina(15,e) and disciplina(6,e)=disciplina(15,e) and disciplina(7,e)=disciplina(15,e) and disciplina(8,e)=disciplina(15,e) and disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=11
								if disciplina(5,e)=disciplina(4,e) then linhas=0
								'if disciplina(5,e)=disciplina(3,e) then linhas=0
								'if disciplina(5,e)=disciplina(2,e) then linhas=0
								'if disciplina(5,e)=disciplina(1,e) then linhas=0
								'if disciplina(5,e)=disciplina(0,e) then linhas=0
							case 6
								linhas=1
								if disciplina(6,e)=disciplina(7,e) then linhas=2
								if disciplina(6,e)=disciplina(8,e) and disciplina(7,e)=disciplina(8,e) then linhas=3
								if disciplina(6,e)=disciplina(9,e) and disciplina(7,e)=disciplina(9,e) and disciplina(8,e)=disciplina(9,e) then linhas=4
								if disciplina(6,e)=disciplina(10,e) and disciplina(7,e)=disciplina(10,e) and disciplina(8,e)=disciplina(10,e) and disciplina(9,e)=disciplina(10,e) then linhas=5
								if disciplina(6,e)=disciplina(11,e) and disciplina(7,e)=disciplina(11,e) and disciplina(8,e)=disciplina(11,e) and disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=6
								if disciplina(6,e)=disciplina(12,e) and disciplina(7,e)=disciplina(12,e) and disciplina(8,e)=disciplina(12,e) and disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=7
								if disciplina(6,e)=disciplina(13,e) and disciplina(7,e)=disciplina(13,e) and disciplina(8,e)=disciplina(13,e) and disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=8
								if disciplina(6,e)=disciplina(14,e) and disciplina(7,e)=disciplina(14,e) and disciplina(8,e)=disciplina(14,e) and disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=9
								if disciplina(6,e)=disciplina(15,e) and disciplina(7,e)=disciplina(15,e) and disciplina(8,e)=disciplina(15,e) and disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=10
								if disciplina(6,e)=disciplina(5,e) then linhas=0
							case 7
								linhas=1
								if disciplina(7,e)=disciplina(8,e) then linhas=2
								if disciplina(7,e)=disciplina(9,e) and disciplina(8,e)=disciplina(9,e) then linhas=3
								if disciplina(7,e)=disciplina(10,e) and disciplina(8,e)=disciplina(10,e) and disciplina(9,e)=disciplina(10,e) then linhas=4
								if disciplina(7,e)=disciplina(11,e) and disciplina(8,e)=disciplina(11,e) and disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=5
								if disciplina(7,e)=disciplina(12,e) and disciplina(8,e)=disciplina(12,e) and disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=6
								if disciplina(7,e)=disciplina(13,e) and disciplina(8,e)=disciplina(13,e) and disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=7
								if disciplina(7,e)=disciplina(14,e) and disciplina(8,e)=disciplina(14,e) and disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=8
								if disciplina(7,e)=disciplina(15,e) and disciplina(8,e)=disciplina(15,e) and disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=9
								if disciplina(7,e)=disciplina(6,e) then linhas=0
							case 8
								linhas=1
								if disciplina(8,e)=disciplina(9,e) then linhas=2
								if disciplina(8,e)=disciplina(10,e) and disciplina(9,e)=disciplina(10,e) then linhas=3
								if disciplina(8,e)=disciplina(11,e) and disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=4
								if disciplina(8,e)=disciplina(12,e) and disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=5
								if disciplina(8,e)=disciplina(13,e) and disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=6
								if disciplina(8,e)=disciplina(14,e) and disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=7
								if disciplina(8,e)=disciplina(15,e) and disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=8
								if disciplina(8,e)=disciplina(7,e) then linhas=0
							case 9
								linhas=1
								if disciplina(9,e)=disciplina(10,e) then linhas=2
								if disciplina(9,e)=disciplina(11,e) and disciplina(10,e)=disciplina(11,e) then linhas=3
								if disciplina(9,e)=disciplina(12,e) and disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=4
								if disciplina(9,e)=disciplina(13,e) and disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=5
								if disciplina(9,e)=disciplina(14,e) and disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=6
								if disciplina(9,e)=disciplina(15,e) and disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=7
								if disciplina(9,e)=disciplina(8,e) then linhas=0
							case 10
								linhas=1
								if disciplina(10,e)=disciplina(11,e) then linhas=2
								if disciplina(10,e)=disciplina(12,e) and disciplina(11,e)=disciplina(12,e) then linhas=3
								if disciplina(10,e)=disciplina(13,e) and disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=4
								if disciplina(10,e)=disciplina(14,e) and disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=5
								if disciplina(10,e)=disciplina(15,e) and disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=6
								if disciplina(10,e)=disciplina(9,e) then linhas=0
							case 11
								linhas=1
								if disciplina(11,e)=disciplina(12,e) then linhas=2
								if disciplina(11,e)=disciplina(13,e) and disciplina(12,e)=disciplina(13,e) then linhas=3
								if disciplina(11,e)=disciplina(14,e) and disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=4
								if disciplina(11,e)=disciplina(15,e) and disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=5
								if disciplina(11,e)=disciplina(10,e) then linhas=0
							case 12
								linhas=1
								if disciplina(12,e)=disciplina(13,e) then linhas=2
								if disciplina(12,e)=disciplina(14,e) and disciplina(13,e)=disciplina(14,e) then linhas=3
								if disciplina(12,e)=disciplina(15,e) and disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=4
								if disciplina(12,e)=disciplina(11,e) then linhas=0
							case 13
								linhas=1
								if disciplina(13,e)=disciplina(14,e) then linhas=2
								if disciplina(13,e)=disciplina(15,e) and disciplina(14,e)=disciplina(15,e) then linhas=3
								if disciplina(13,e)=disciplina(12,e) then linhas=0
							case 14
								linhas=1
								if disciplina(14,e)=disciplina(15,e) then linhas=2
								if disciplina(14,e)=disciplina(13,e) then linhas=0
							case 15
								linhas=1
								if disciplina(15,e)=disciplina(14,e) then linhas=0
						end select

					if linhas>0 then
%>					
	<td class="campor" valign=top style="background-color: <%=backfundo(d,e)%>" rowspan="<%=linhas%>"><%=disciplina(d,e)%><%=compl(d,e)%>
	</td>
<%
					end if
coluna=coluna+1
'leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & planilha(d,e) & """"
'leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = objXL.Cells(" & linha & ", " & coluna & ").Value & chr(10) & """ & planilha2(d,e) & """"
leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = """ & planilha(d,e) & """ & chr(10) & """ & planilha2(d,e) & """"
'leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = objXL.Cells(" & linha & ", " & coluna & ").Value & """ & planilha2(d,e) & """"
if planilha3(d,e)<>"" or planilha(d,e)<>"0" then
	leitura.writeline "objXL.Cells(" & linha & ", " & coluna & ").Value = objXL.Cells(" & linha & ", " & coluna & ").Value & chr(10) & """ & planilha3(d,e) & """"
end if
					next 'sturma e
linha=linha+1
%>
</tr>
<%
				next 'horainicio d
				'erase pula

			'next 'diasemana c
%>
</table>
<br>
<%
'		next 'turnograde b
		'erase turnograde
		'erase pula
	
'	next 'periodoletivo a
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