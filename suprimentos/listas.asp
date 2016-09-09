<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.redirect "../intranet.asp"
if session("a94")="N" or session("a94")="" then response.redirect "../intranet.asp"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Relatórios - Suprimentos</title>
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

Sub ImprimeQuadro(titulo)
'*************** inicio teste **********************
response.write "Relatório de: " & titulo
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
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
response.write "<p style='margin-top:0;margin-bottom:0'>"
'*************** fim teste **********************
nomefile="suprimentos_" & titulo & ".vbs"
response.write "<a href='../temp/" & nomefile & "'>"
response.write "<img src='../images/msexcel.gif' width='16' height='16' border='0' alt=''>"
response.write "</a>"
End Sub

Sub Planilha(titulo)
	rs.movefirst
	caminho="c:\inetpub\wwwroot\rh\temp\" 
	a=session.sessionid
	nomefile="suprimentos_" & titulo & ".vbs"
	lote=caminho & nomefile
	Set leitura=arquivo.CreateTextFile(lote, true)

	leitura.writeline "Dim objXL"
	leitura.writeline "Set objXL = WScript.CreateObject(""Excel.Application"")"
	leitura.writeline "objXL.Visible = TRUE"
	leitura.writeline "objXL.WorkBooks.Add"
	leitura.writeline "objXL.Cells(1, 1).Value = ""Suprimentos - Centro Universitário UNIFIEO"""
	leitura.writeline "objXL.Cells(3, 1).Value = ""Relatorio: """
	leitura.writeline "objXL.Cells(3, 2).Value = """ & titulo & """"

	for a= 0 to rs.fields.count-1
		leitura.writeline "objXL.Cells(5," & a+1 &").Value = """ & ucase(rs.fields(a).name) & """"
	next
	ultcol=chr(64+rs.fields.count)
	if ultcol="A" then ultcol="B"
	leitura.writeline "objXL.Range(""A5:" & ultcol & "5"").Select"
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
			for a= 0 to rs.fields.count-1
				if rs.fields(a).type=7 then
					leitura.writeline "objXL.Cells(" & linha & ", " & a+1 & ").Value = """ & dtaccess(rs.fields(a)) & """"
				elseif rs.fields(a).type=5 then
					leitura.writeline "objXL.Cells(" & linha & ", " & a+1 & ").Value = """ & "=" & nraccess(rs.fields(a)) & """"
				else
					leitura.writeline "objXL.Cells(" & linha & ", " & a+1 & ").Value = """ & rs.fields(a) & """"
				end if
				'texto-202 data-7 inteiro-2 duplo-5 inteiro?-3
			next
		linha=linha+1
		rs.movenext:	loop

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
End Sub

dim conexao, rs, rs2, arquivo, leitura
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
Set arquivo=CreateObject("Scripting.FileSystemObject")

if request.form("opcao")="" or request.form("pesquisar")="" then
%>
<!-- -->
<form method="POST" action="listas.asp" name="form">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo colspan=2><p style="margin-top:0;margin-bottom:0;color:Blue;font-size:10pt;text-align:left">
<b>Opções para emissão de relatório</font></p></td></tr>
<%a=request.form("opcao"):b="checked"%>
<%i="categoria"%>
	<tr><td class=titulo><input type="radio" name="opcao" onClick="javascript:submit()" value="<%=i%>" <%if a=i then response.write b%>> Categorias</td></tr>
<%i="categoria_itens"%>
	<tr><td class=titulo><input type="radio" name="opcao" onClick="javascript:submit()" value="<%=i%>" <%if a=i then response.write b%>> Categorias com itens
<%	if request.form("opcao")=i then
		response.write "<br><font color='blue'>&nbsp;Categorias: <select class='a' size='1' name='sel_cat_itens' >"
		sql1="select distinct c.id_cat, c.descricao from uniforme_categoria c, uniforme_link l where c.id_cat=l.id_cat order by c.descricao"
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		response.write "<option value=''>Todas categorias...</option>"
		rs.movefirst:do while not rs.eof
		if request.form("sel_cat_itens")=rs("id_cat") then temp="selected" else temp=""
		response.write "<option value='" & rs("id_cat") & "' " & temp & ">" & rs("descricao") & "</option>"
		rs.movenext:loop:rs.close
		response.write "</select>"
		response.write "<br>&nbsp;<input type='checkbox' name='sel_saldo' value='ON'> <font color='blue'>mostrar estoque atual"
	end if	
%>
</td></tr>
<%i="estoque_saldo"%>
	<tr><td class=titulo><input type="radio" name="opcao" onClick="javascript:submit()"  value="<%=i%>" <%if a=i then response.write b%>> Estoque (apenas saldo)
<%	if request.form("opcao")=i then
		response.write "<br><font color='blue'>&nbsp;Categorias: <select class='a' size='1' name='sel_cat_itens0' >"
		sql1="select distinct c.id_cat, c.descricao from uniforme_categoria c, uniforme_link l where c.id_cat=l.id_cat order by c.descricao"
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		response.write "<option value=''>Todas categorias...</option>"
		rs.movefirst:do while not rs.eof
		if request.form("sel_cat_itens0")=rs("id_cat") then temp="selected" else temp=""
		response.write "<option value='" & rs("id_cat") & "' " & temp & ">" & rs("descricao") & "</option>"
		rs.movenext:loop:rs.close
		response.write "</select>"
		response.write "<br>&nbsp;Ordem: <select class='a' size='1' name='sel_ordem0'>"
		response.write "<option value='PC'>Itens agrupados por categoria</option>"
		response.write "<option value='OA'>Todos itens em ordem alfabética</option>"
		response.write "</select>"
	end if	
%>
	</td></tr>
<%i="estoque_mov"%>
	<tr><td class=titulo><input type="radio" name="opcao" onClick="javascript:submit()"  value="<%=i%>" <%if a=i then response.write b%>> Estoque (com movimento)
<%	if request.form("opcao")=i then
		response.write "<br><font color='blue'>&nbsp;Categorias: <select class='a' size='1' name='sel_cat_itens' onChange='javascript:submit()'>"
		sql1="select distinct c.id_cat, c.descricao from uniforme_categoria c, uniforme_link l where c.id_cat=l.id_cat order by c.descricao"
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		response.write "<option value=''>Todas categorias...</option>"
		rs.movefirst:do while not rs.eof
		if request.form("sel_cat_itens")<>"" then sel_cat_itens=cint(request.form("sel_cat_itens")) else sel_cat_itens=0
		if sel_cat_itens=cint(rs("id_cat")) then temp="selected" else temp=""
		response.write "<option value='" & rs("id_cat") & "' " & temp & ">" & rs("descricao") & "</option>"
		rs.movenext:loop:rs.close
		response.write "</select>"
			if request.form("sel_cat_itens")="" then sel_cat=">=0" else sel_cat="=" & request.form("sel_cat_itens")
		response.write "<br><font color='blue'>&nbsp;Items: <select class='a' size='1' name='sel_itens' >"
		sql1="SELECT i.id_item, i.descricao, i.sequencia, i.tamanho FROM uniforme_item AS i INNER JOIN uniforme_link AS l ON i.id_item = l.id_item " & _
		"WHERE l.id_cat" & sel_cat & " ORDER BY i.descricao, i.sequencia "
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		response.write "<option value=''>Todos itens...</option>"
		rs.movefirst:do while not rs.eof
		if request.form("sel_itens")=rs("id_item") then temp="selected" else temp=""
		response.write "<option value='" & rs("id_item") & "' " & temp & ">" & rs("descricao") & " (" & rs("tamanho") & ")" & "</option>"
		rs.movenext:loop:rs.close
		response.write "</select>"
		if request.form("datai")="" then datai=dateserial(year(now),month(now),1) else datai=request.form("datai")
		if request.form("dataf")="" then dataf=dateserial(year(now),month(now)+1,1)-1 else dataf=request.form("dataf")
		response.write "<br><font color='blue'>&nbsp;Periodo: " 
		response.write "<input type=text name=datai value='" & datai & "' size=8>"
		response.write " até "
		response.write "<input type=text name=dataf value='" & dataf & "' size=8>"
	end if	
%>
	</td></tr>
	<tr><td class=titulo><hr></td></tr>
<%i="an_func_sem_uniforme"%>
	<tr><td class=titulo><input type="radio" name="opcao" onClick="javascript:submit()"  value="<%=i%>" <%if a=i then response.write b%>> Funcionários sem uniforme
<%	if request.form("opcao")=i then
		if request.form("ultimo_ent")="" then ultimo_ent=12 else ultimo_ent=request.form("ultimo_ent")
		response.write "<br><font color='blue'>&nbsp;Não recebeu uniforme nos últimos <input type=text class=a name=ultimo_ent value=" & ultimo_ent & " size=2> meses."
		response.write "<br>&nbsp;Ordem: <select class='a' size='1' name='sel_ordem2'>"
		response.write "<option value='oC'>ordem de Categoria</option>"
		response.write "<option value='oS'>ordem de Setor</option>"
		response.write "</select>"
	end if	
%>
	</td></tr>
<%i="an_func_com_usado"%>
	<tr><td class=titulo><input type="radio" name="opcao" onClick="javascript:submit()"  value="<%=i%>" <%if a=i then response.write b%>> Troca de uniforme usado
<%	if request.form("opcao")=i then
		if request.form("usado_ent")="" then usado_ent=6 else usado_ent=request.form("usado_ent")
		response.write "<br><font color='blue'>&nbsp;Entregue nos últimos <input type=text class=a name=usado_ent value=" & usado_ent & " size=2> meses."
		response.write "<br>&nbsp;Ordem: <select class='a' size='1' name='sel_ordem1'>"
		response.write "<option value='oC'>ordem de Categoria</option>"
		response.write "<option value='oS'>ordem de Setor</option>"
		response.write "</select>"
	end if	
%>
	</td></tr>
<%i="an_total_setor"%>
	<tr><td class=titulo><input type="radio" name="opcao" onClick="javascript:submit()"  value="<%=i%>" <%if a=i then response.write b%>> Total por departamento
<%	if request.form("opcao")=i then
		if request.form("datai2")="" then datai2=dateserial(year(now),month(now),1) else datai2=request.form("datai2")
		if request.form("dataf2")="" then dataf2=dateserial(year(now),month(now)+1,1)-1 else dataf2=request.form("dataf2")
		response.write "<br><font color='blue'>&nbsp;Periodo: " 
		response.write "<input class=a type=text name=datai2 value='" & datai2 & "' size=8>"
		response.write " até "
		response.write "<input class=a type=text name=dataf2 value='" & dataf2 & "' size=8>"
	end if	
%>
	</td></tr>

<tr>
	<td class=fundo align="center"><hr>
	<input type="submit" value="Pesquisar" class="button" name="pesquisar" onfocus="javascript:window.status='Clique aqui para pesquisar'">
	</td>
</tr>
</table>
</form>
<%
end if

if request.form("pesquisar")<>"" then
'response.write request.form
'*************** totais *********************
database=cdate(request.form("database"))

if request.form("opcao")="categoria" then
	titulo="Categorias"
	sql1="select descricao from uniforme_categoria order by descricao"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	ImprimeQuadro(titulo)
	Planilha(titulo)
end if 'categoria

if request.form("opcao")="categoria_itens" then
	titulo="Categorias com items"
	if request.form("sel_cat_itens")<>"" then selecao1="c.id_cat=" & request.form("sel_cat_itens") & " " else selecao1="c.id_cat>0 "
	if request.form("sel_saldo")="ON" then selecao2=" s.tnovo, s.tusado," else selecao2=""
	sql1="SELECT c.descricao AS categoria, i.descricao AS item, i.tamanho, i.codigoRM," & selecao2 & " i.preco " & _
	"FROM ((uniforme_link l INNER JOIN uniforme_categoria c ON l.id_cat = c.id_cat) " & _
	"INNER JOIN uniforme_item i ON l.id_item = i.id_item) LEFT JOIN uniforme_saldo s ON i.id_item = s.id_item " & _
	"WHERE " & selecao1 & _
	"ORDER BY c.descricao, i.descricao, i.sequencia "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	ImprimeQuadro(titulo)
	Planilha(titulo)
end if 'categoria_itens

if request.form("opcao")="estoque_saldo" then
	titulo="Saldo do Estoque de Uniformes"
	if request.form("sel_cat_itens0")<>"" then selecao1="c.id_cat=" & request.form("sel_cat_itens0") & " " else selecao1="c.id_cat>0 "
	if request.form("sel_ordem0")="PC" then ordem="c.descricao, i.descricao ":inicio="c.descricao as categoria, "
	if request.form("sel_ordem0")="OA" then ordem="i.descricao ":inicio=""
	sql1="SELECT distinct " & inicio & "i.descricao AS item, i.tamanho, i.codigoRM, s.tnovo, s.tusado, i.preco " & _
	"FROM ((uniforme_link l INNER JOIN uniforme_categoria c ON l.id_cat = c.id_cat) " & _
	"INNER JOIN uniforme_item i ON l.id_item = i.id_item) LEFT JOIN uniforme_saldo s ON i.id_item = s.id_item " & _
	"WHERE " & selecao1 & _
	"ORDER BY " & ordem
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	ImprimeQuadro(titulo)
	Planilha(titulo)
end if 'estoque_saldo

if request.form("opcao")="estoque_mov" then
	titulo="Saldo do Estoque com movimentação dos Uniformes"
	if request.form("sel_cat_itens")<>"" then selecao1="l.id_cat=" & request.form("sel_cat_itens") & " " else selecao1="l.id_cat>0 "
	if request.form("sel_itens")<>"" then selecao2="i.id_item=" & request.form("sel_itens") & " " else selecao2="i.id_item>0 "
	datai=request.form("datai")
	dataf=request.form("dataf")
	sql1="select * from (" & _
	"SELECT distinct [e].id_item, i.descricao AS desc_item, i.tamanho, i.preco, min(convert(datetime,'" & dtaccess(datai) & "')-1) AS datamov, min('Saldo Anterior') AS desc_mov, " & _
	"Sum(e.qt_novo*[tipo])+Min(i.qt_novo) AS novo, Sum(e.qt_usado*[tipo])+Min(i.qt_usado) AS usado, Null AS chapa " & _
	"FROM uniforme_tpmov AS tm INNER JOIN ((uniforme_item AS i INNER JOIN uniforme_estoque AS e ON i.id_item=e.id_item) INNER JOIN uniforme_link AS l ON i.id_item=l.id_item) ON tm.id_mov=e.id_mov " & _
	"WHERE e.dt_movimento<'" & dtaccess(datai) & "' and " & selecao1 & " and " & selecao2 & " " & _
	"GROUP BY l.id_cat, e.id_item, i.descricao, i.tamanho, i.preco " & _
	"union all " & _
	"SELECT distinct e.id_item, i.descricao as desc_item, i.tamanho, i.preco, e.dt_movimento, tm.descricao as desc_mov, (e.qt_novo*tm.tipo) as novo, (e.qt_usado*tm.tipo) as usado, f.nome " & _
	"FROM (uniforme_tpmov AS tm INNER JOIN ((uniforme_item AS i INNER JOIN uniforme_estoque AS e ON i.id_item=e.id_item) INNER JOIN uniforme_link AS l ON i.id_item=l.id_item) ON tm.id_mov=e.id_mov) LEFT JOIN corporerm.dbo.PFUNC AS f ON e.chapa=f.CHAPA collate database_default " & _
	"WHERE e.dt_movimento between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' and " & selecao1 & " and " & selecao2 & " " & _
	") as l order by desc_item, tamanho, datamov "
	'response.write "<br><br>" & sql1
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	ImprimeQuadro(titulo)
	Planilha(titulo)
end if

if request.form("opcao")="an_func_com_usado" then
	titulo="Uniformes usados entregue à Funcionários"
	meses=request.form("usado_ent")
	iniciorel=dateserial(year(now),month(now)-meses,day(now))
	if request.form("sel_ordem1")="oC" then ordem="c.descricao, f.nome ":inicio=""
	if request.form("sel_ordem1")="oS" then ordem="f.secao, f.nome ":inicio=""

	sql1="SELECT c.descricao AS desc_cat, [e].chapa, f.NOME, f.Secao, [e].dt_movimento, i.descricao AS desc_item, i.tamanho, e.qt_usado " & _
	"FROM (((uniforme_item AS i INNER JOIN uniforme_estoque AS e ON i.id_item=e.id_item) INNER JOIN uniforme_func_cat AS fc ON e.chapa=fc.chapa) " & _
	"INNER JOIN uniforme_categoria AS c ON fc.id_cat=c.id_cat) INNER JOIN qry_funcionarios AS f ON e.chapa=f.CHAPA collate database_default " & _
	"WHERE fc.id_cat>0 AND e.id_item>0 AND e.qt_usado>0 AND e.id_mov=1 AND f.CODSITUACAO<>'D' and e.dt_movimento>='" & dtaccess(iniciorel) & "' " & _
	"ORDER by " & ordem
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	ImprimeQuadro(titulo)
	Planilha(titulo)
end if

if request.form("opcao")="an_func_sem_uniforme" then
	titulo="Funcionários sem uniformes entregues"
	meses=request.form("ultimo_ent")
	iniciorel=dateserial(year(now),month(now)-meses,day(now))
	if request.form("sel_ordem2")="oC" then ordem="c.descricao, f.nome ":inicio=""
	if request.form("sel_ordem2")="oS" then ordem="f.secao, f.nome ":inicio=""
	sql1="SELECT fc.chapa, f.NOME, f.Secao, c.descricao AS desc_cat " & _
	"FROM (uniforme_func_cat AS fc INNER JOIN qry_funcionarios AS f ON fc.chapa=f.CHAPA collate database_default) INNER JOIN uniforme_categoria AS c ON fc.id_cat=c.id_cat " & _
	"WHERE fc.chapa Not In (SELECT distinct chapa FROM uniforme_estoque WHERE dt_movimento>='" & dtaccess(iniciorel) & "' AND id_mov=1) " & _
	"AND fc.id_cat<>8 AND f.CODSITUACAO In ('A','F','Z') " & _
	"ORDER by " & ordem
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	ImprimeQuadro(titulo)
	Planilha(titulo)
end if

if request.form("opcao")="an_total_setor" then
	titulo="Total em valores do Uniformes entregues"
	datai2=request.form("datai2")
	dataf2=request.form("dataf2")
	sql1="SELECT f.Secao, i.descricao AS Desc_uniforme, i.tamanho, i.preco, " & _
	"Entregue  =sum(case when tp.id_mov=1 then e.qt_novo+e.qt_usado else 0 end), " & _
	"Devolvido =sum(case when tp.id_mov=3 then e.qt_novo+e.qt_usado else 0 end), " & _
	"Utilizados=sum(case when tp.id_mov=1 then e.qt_novo+e.qt_usado else 0 end - case when tp.id_mov=3 then e.qt_novo+e.qt_usado else 0 end), " & _
	"Total     =sum(case when tp.id_mov=1 then e.qt_novo+e.qt_usado else 0 end - case when tp.id_mov=3 then e.qt_novo+e.qt_usado else 0 end)*preco " & _
	"FROM (uniforme_tpmov AS tp INNER JOIN (uniforme_item AS i INNER JOIN uniforme_estoque AS e ON i.id_item=e.id_item) ON tp.id_mov=e.id_mov) INNER JOIN qry_funcionarios AS f ON e.chapa=f.CHAPA collate database_default " & _
	"WHERE e.dt_movimento Between '" &  dtaccess(datai2) & "' And '" &  dtaccess(dataf2) & "' AND e.id_item>0 AND e.id_mov Not In (4) " & _
	"GROUP BY f.CODSECAO, f.Secao, i.descricao, i.tamanho, i.sequencia, i.preco " & _
	"HAVING f.CODSECAO>'0' " & _
	"ORDER BY f.CODSECAO, i.descricao, i.sequencia "
	'response.write "<br><br>" & sql1
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	ImprimeQuadro(titulo)
	Planilha(titulo)
end if

%>
<p>
<%
end if
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
set leitura=nothing
set arquivo=nothing

%>

<!-- -->
</body>
</html>