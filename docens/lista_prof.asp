<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.redirect "../intranet.asp"
if session("a97")="N" or session("a97")="" then response.redirect "../intranet.asp"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Lista Personalizada</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head> 
<body>
<%
inicio=now()
Sub CentralizaCelulas
	leitura.writeline "objXL.Selection.HorizontalAlignment = -4108" 'xlCenter"
	leitura.writeline "objXL.Selection.VerticalAlignment = -4107" 'xlBottom"
	leitura.writeline "objXL.Selection.WrapText = False"
	leitura.writeline "objXL.Selection.Orientation = 0"
	leitura.writeline "objXL.Selection.AddIndent = False"
	leitura.writeline "objXL.Selection.ShrinkToFit = False"
	leitura.writeline "objXL.Selection.MergeCells = False"
	leitura.writeline "objXL.Selection.Merge"
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

dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form="" then
%>
<!-- -->
<form method="POST" action="lista_prof.asp" name="form">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo colspan=2><p style="margin-top:0;margin-bottom:0;color:Blue;font-size:10pt;text-align:left">
<b>Seleção para lista personalizada</font></p>
	</td></tr>
<tr><td class=titulo><hr>Curso</td></tr>
<tr><td class=fundo>
	<select size="1" name="curso" onfocus="javascript:window.status='Selecione o curso'" >
<%
sql2="select g.coddoc, c.curso from g2ch g, g2cursoeve c where g.coddoc=c.coddoc group by  g.coddoc, c.curso ORDER BY c.CURSO "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value='all'>Todos cursos</option>"
rs.movefirst:do while not rs.eof
if chapa=rs("coddoc") then tempc="selected" else tempc=""
%>
		<option value="<%=rs("coddoc")%>" <%=tempc%>><%=rs("curso")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select></td></tr>
<tr><td class=titulo>Campos</td></tr>
<tr><td class=fundo>
	<input type="checkbox" name="titulacao" value="on">Titul.Real
	<input type="checkbox" name="titulacaomec" value="on">Titul.MEC
	<input type="checkbox" name="nascimento" value="on">Data Nasc.
	<input type="checkbox" name="aniversario" value="on">Aniversário
	</td>
</tr>
<tr>
	<td class=fundo>
	<input type="checkbox" name="endereco" value="on">Endereço
	<input type="checkbox" name="email" value="on">Email
	<input type="checkbox" name="telefone" value="on">Telefones
	<input type="checkbox" name="admissao" value="on">Data Admissão
	</td>
</tr>
<tr>
	<td class=fundo><!-- -->
	<input type="checkbox" name="aulas" value="on">nº Aulas
	<input type="checkbox" name="atividades" value="on">nº Ativ.<!-- -->
	<input type="checkbox" name="total" value="on">Total CH
	<input type="checkbox" name="sexo" value="on">Sexo
	<input type="checkbox" name="cpf" value="on">CPF
	</td>
</tr>
<tr>
	<td class=fundo><!--
	<input type="checkbox" name="formacao" value="on">Curso Formação-->
	<input type="checkbox" name="setor" value="on">Setor
</tr>
<tr>
	<td class=fundo>
	Data-base das informações: <input type="text" name="database" size="8" value="<%=formatdatetime(now(),2)%>">
	<br>Incluir Demitidos/Afastados: <input type="checkbox" name="demitidos" value="ON">
                                  
</tr>
<tr>
	<td class=fundo align="center"><hr>
	<input type="submit" value="Pesquisar" class="button" name="pesquisar" onfocus="javascript:window.status='Clique aqui para pesquisar'">
	</td>
</tr>
</table>
</form>
<%
end if

if request.form<>"" then
	if request.form("titulacao")   ="on" then titulacao   =1 else titulacao   =0
	if request.form("titulacaomec")="on" then titulacaomec=1 else titulacaomec=0
	if request.form("nascimento")  ="on" then nascimento  =1 else nascimento  =0
	if request.form("aniversario") ="on" then aniversario =1 else aniversario =0
	if request.form("endereco")    ="on" then endereco    =1 else endereco    =0
	if request.form("email")       ="on" then email       =1 else email       =0
	if request.form("telefone")    ="on" then telefone    =1 else telefone    =0
	if request.form("admissao")    ="on" then admissao    =1 else admissao    =0
	if request.form("aulas")       ="on" then aulas       =1 else aulas       =0
	if request.form("atividades")  ="on" then atividades  =1 else atividades  =0
	if request.form("total")       ="on" then total       =1 else total       =0
	if request.form("sexo")        ="on" then sexo        =1 else sexo        =0
	if request.form("cpf")         ="on" then cpf         =1 else cpf         =0
'	if request.form("formacao")    ="on" then formacao    =1 else formacao    =0
	if request.form("setor")       ="on" then setor       =1 else setor       =0
	if request.form("demitidos")="ON" then incdem=1 else incdem=0
'*************** totais *********************
database=cdate(now)
database=cdate(request.form("database"))
sql1="delete from ttcargahoraria_ch where sessao='" & session("usuariomaster") & "' ":conexao.execute sql1
sql2="INSERT INTO ttcargahoraria_ch ( sessao, tipoch, CHAPA, cargahoraria, [database] ) SELECT '" & session("usuariomaster") & "', 1 , chapa1, Sum(ta), '" & dtaccess(database) & "' FROM g2ch WHERE '" & dtaccess(database) & "' Between [inicio] And [termino] GROUP BY chapa1 "
sql3="INSERT INTO ttcargahoraria_ch ( sessao, tipoch, CHAPA, cargahoraria, [database] ) SELECT '" & session("usuariomaster") & "', 2 , CHAPA, Sum(case when codeve is null or codeve='' then 0 else CH end), '" & dtaccess(database) & "' FROM n_indicacoes WHERE '" & dtaccess(database) & "' Between [mand_ini] And [mand_fim] GROUP BY CHAPA "
sql4="INSERT INTO ttcargahoraria_ch ( sessao, tipoch, CHAPA, cargahoraria, [database] ) SELECT '" & session("usuariomaster") & "', 3 , CHAPA, Sum(CH), '" & dtaccess(database) & "' FROM grades_rt WHERE '" & dtaccess(database) & "' Between [inicio] And [fim] GROUP BY CHAPA "
sql5="INSERT INTO ttcargahoraria_ch ( sessao, tipoch, CHAPA, cargahoraria, [database] ) SELECT sessao, 4 , CHAPA, Sum(cargahoraria) AS SomaDecargahoraria, '" & dtaccess(database) & "' FROM ttcargahoraria_ch GROUP BY sessao, CHAPA, [database] HAVING sessao='" & session("usuariomaster") & "' and [database]='" & dtaccess(database) & "' "
sql12="SELECT sessao, [database] FROM ttcargahoraria_ch GROUP BY sessao, [database] HAVING sessao='" & session("usuariomaster") & "' and [database]='" & dtaccess(database) & "' "
rs.Open sql12, ,adOpenStatic, adLockReadOnly
if rs.recordcount=0 then
	conexao.execute sql1
	conexao.execute sql2:conexao.execute sql3:conexao.execute sql4:conexao.execute sql5
end if
rs.close

sqlp=sqlp & "from pfunc f, ppessoa p, pcodinstrucao i " & _
"where f.codpessoa=p.codigo and p.grauinstrucao=i.codcliente " & _
"and f.codsituacao in ('A','F','Z') and f.codsindicato='03' "

sqlp="select f.chapa, f.nome "
if titulacao   =1 then sqlp=sqlp & ",instrucao as titulacaoreal "
if titulacaomec=1 then sqlp=sqlp & ",instrucaomec as titulacaomec "
if nascimento  =1 then sqlp=sqlp & ",dtnascimento as nascimento "
if aniversario =1 then sqlp=sqlp & ",day(dtnascimento) as dia_aniv ,month(dtnascimento) as mes_aniv "
if sexo        =1 then sqlp=sqlp & ",sexo "
if endereco    =1 then sqlp=sqlp & ",f.rua ,f.numero ,f.complemento ,f.bairro ,f.cidade ,f.cep "
if email       =1 then sqlp=sqlp & ",f.email "
if telefone    =1 then sqlp=sqlp & ",telefone1 ,telefone2 ,telefone3 ,fax "
if admissao    =1 then sqlp=sqlp & ",dataadmissao as admissao "
if aulas       =1 then sqlp=sqlp & ",aulas=sum(case when tipoch=1 then cargahoraria else 0 end) "
if atividades  =1 then sqlp=sqlp & ",atividades=sum(case when tipoch=2 Or tipoch=3 then [cargahoraria] else 0 end ) "
if total       =1 then sqlp=sqlp & ",total=sum(case when [tipoch]=4 then [cargahoraria] else 0 end) "
if cpf         =1 then sqlp=sqlp & ",cpf "
'if formacao    =1 then sqlp=sqlp & ",u.curso + ' (' + u.instituicao + ' / ' + u.anoconclusao + ')' as formacao "
if setor       =1 then sqlp=sqlp & ",s.descricao as setor "
'u.CURSO

if incdem=1 then
	filtrodem=""
else
	filtrodem=" AND f.codsituacao in ('A','F','E','Z') "
end if

if request.form("curso")<>"all" then
	sqlp=sqlp & "from (((dc_professor f " & _
	"LEFT JOIN ttcargahoraria_ch t ON f.chapa collate database_default=t.chapa) LEFT JOIN " & _
	"(SELECT c.chapa1, c.coddoc FROM g2ch c " & _
	"WHERE '" & dtaccess(database) & "' Between [inicio] And [termino] GROUP BY c.chapa1, c.coddoc) c ON f.chapa collate database_default=c.chapa1) " & _
	"/*LEFT JOIN UPROFFORMACAO_ u ON (f.GRAUINSTRUCAO collate database_default=u.CODINSTRUCAO) AND (f.CHAPA collate database_default=u.CODPROF)*/ ) " & _
	"LEFT JOIN corporerm.dbo.PSECAO S ON (S.CODIGO collate database_default=F.CODSECAO) " & _
	"WHERE /*f.codsindicato='03' AND*/ t.sessao='" & session("usuariomaster") & "' " & filtrodem & " " & _
	"AND t.[database]='" & dtaccess(database) & "' "
	sqlp=sqlp & "AND c.coddoc='" & request.form("curso") & "' "
	sql2="select curso from g2cursoeve where coddoc='" & request.form("curso") & "' "
	rs.Open sql2, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then nomecurso=rs("curso") else nomecurso=""
	rs.close
	'"WHERE f.codsituacao In ('A','F','Z') AND f.codsindicato='03' AND t.sessao='" & session("usuariomaster") & "' " & _
else
	sqlp=sqlp & "from ((dc_professor f " & _
	"LEFT JOIN ttcargahoraria_ch t ON f.chapa collate database_default=t.chapa) " & _
	"/*LEFT JOIN UPROFFORMACAO_ u ON (f.GRAUINSTRUCAO collate database_default=u.CODINSTRUCAO) AND (f.CHAPA collate database_default=u.CODPROF)*/ ) " & _
	"LEFT JOIN corporerm.dbo.PSECAO S ON (S.CODIGO collate database_default=F.CODSECAO) " & _
	"WHERE /*f.codsindicato='03' AND*/ t.sessao='" & session("usuariomaster") & "' " & filtrodem & " " & _
	"AND t.[database]='" & dtaccess(database) & "' "
	nomecurso="Todos"
	'"WHERE f.codsituacao In ('A','F','Z') AND f.codsindicato='03' AND t.sessao='" & session("usuariomaster") & "' " & _
end if
sqlp=sqlp & "GROUP BY f.chapa, f.nome "
if titulacao   =1 then sqlp=sqlp & ",instrucao "
if titulacaomec=1 then sqlp=sqlp & ",instrucaomec "
if nascimento  =1 then sqlp=sqlp & ",dtnascimento "
if aniversario =1 then sqlp=sqlp & ",day(dtnascimento) ,month(dtnascimento) "
if sexo        =1 then sqlp=sqlp & ",sexo "
if endereco    =1 then sqlp=sqlp & ",f.rua ,f.numero ,f.complemento ,f.bairro ,f.cidade ,f.cep "
if email       =1 then sqlp=sqlp & ",f.email "
if telefone    =1 then sqlp=sqlp & ",telefone1 ,telefone2 ,telefone3 ,fax "
if admissao    =1 then sqlp=sqlp & ",dataadmissao "
if aulas       =1 then sqlp=sqlp & " "
if atividades  =1 then sqlp=sqlp & " "
if total       =1 then sqlp=sqlp & " "
if cpf         =1 then sqlp=sqlp & ",cpf "
'if formacao    =1 then sqlp=sqlp & ",u.curso + ' (' + u.instituicao + ' / ' + u.anoconclusao + ')' "
if setor       =1 then sqlp=sqlp & ",s.descricao "
'response.write sqlp
rs.Open sqlp, ,adOpenStatic, adLockReadOnly
'*************** inicio teste **********************
response.write "Curso: " & nomecurso
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=""titulor"">" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=""campor"" nowrap>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
response.write "<p>"
'*************** fim teste **********************
rs.movefirst

caminho="c:\inetpub\wwwroot\rh\temp\"
nomefile="lista_prof_" & session.sessionid & ".vbs"
lote=caminho & nomefile
Set arquivo=CreateObject("Scripting.FileSystemObject")
Set leitura=arquivo.CreateTextFile(lote, true)
%>
<%
leitura.writeline "Dim objXL"
leitura.writeline "Set objXL = WScript.CreateObject(""Excel.Application"")"
leitura.writeline "objXL.Visible = TRUE"
leitura.writeline "objXL.WorkBooks.Add"
leitura.writeline "objXL.Cells(1, 1).Value = ""Lista de Professores - Centro Universitário UNIFIEO"""
leitura.writeline "objXL.Cells(3, 1).Value = ""Curso: """
leitura.writeline "objXL.Cells(3, 2).Value = """ & nomecurso & """"

for a= 0 to rs.fields.count-1
	leitura.writeline "objXL.Cells(5," & a+1 &").Value = """ & ucase(rs.fields(a).name) & """"
next

leitura.writeline "objXL.Range(""A5:x5"").Select"
Bordas
leitura.writeline "objXL.Selection.Font.Bold = True"
leitura.writeline "objXL.Selection.Interior.ColorIndex = 15" 'gray
leitura.writeline "objXL.Selection.Interior.Pattern = 1 " 'xlSolid
leitura.writeline "objXL.Selection.Font.ColorIndex = 1" 'black

leitura.writeline "objXL.Range(""A1:G1"").Select"
leitura.writeline "objXL.Selection.Font.Bold = True"
CentralizaCelulas
leitura.writeline "objXL.Rows(""5:5"").RowHeight = 23"
leitura.writeline "objXL.Rows(""5:5"").Select"
leitura.writeline "objXL.Selection.HorizontalAlignment = -4108" 'xlCenter
leitura.writeline "objXL.Selection.VerticalAlignment = -4160" 'xlTop
'leitura.writeline "objXL.Selection.WrapText = True"
leitura.writeline "objXL.Selection.Orientation = 0"
leitura.writeline "objXL.Selection.Font.Name = ""Arial"""
leitura.writeline "objXL.Selection.Font.Size = 8"
leitura.writeline "objXL.Columns(1).ColumnWidth = 6"
leitura.writeline "objXL.Columns(2).ColumnWidth = 50"
leitura.writeline "objXL.Range(""A3:A3"").Select"
CentralizaCelulas
Esquerda	
leitura.writeline "objXL.Range(""A4:A4"").Select"
CentralizaCelulas
Esquerda	
linha=6
	rs.movefirst
	do while not rs.eof 
		'response.write "<br>"
		for a= 0 to rs.fields.count-1
			'response.write rs.fields(a) & " >> " & rs.fields(a).type & " | "
			if rs.fields(a).type=7 or rs.fields(a).type=135 then
				texto=day(rs.fields(a))&"/"&month(rs.fields(a))&"/"&year(rs.fields(a))
				leitura.writeline "objXL.Cells(" & linha & ", " & a+1 & ").Value = """ & "" & dtaccess(rs.fields(a)) & """"
			'elseif rs.fields(a).type=5 or rs.fields(a).type=2 then
			'	leitura.writeline "objXL.Cells(" & linha & ", " & a+1 & ").Value = """ & "=" & nraccess(rs.fields(a)) & """"
			else
				leitura.writeline "objXL.Cells(" & linha & ", " & a+1 & ").Value = """ & rs.fields(a) & """"
			end if
			'texto-202 data-7 inteiro-2 duplo-5
			'texto-200 data-135 inteiro-2 duplo-5
		next
	linha=linha+1
	rs.movenext
	loop

	rs.close
	leitura.writeline "objXL.Range(""A6:AB" & linha-1 & """).Select"
	Bordas
	leitura.writeline "objXL.Cells.Select"
	leitura.writeline "objXL.Selection.Font.Name = ""Tahoma"""
	leitura.writeline "objXL.Selection.Font.Size = 8"
	leitura.writeline "objXL.Cells.EntireColumn.AutoFit"
	leitura.writeline "objXL.Range(""A2"").Select"
	leitura.writeline "objXL.Columns(5).ColumnWidth = 10"
	leitura.close
	set leitura=nothing
	set arquivo=nothing
%>
<p>
Selecione o quadro acima e cole em uma planilha ou abra uma <a href="../temp/<%=nomefile%>">Planilha Excel</a> automaticamente.
<%
end if
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>

<!-- -->
</body>
</html>