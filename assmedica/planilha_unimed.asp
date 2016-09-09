<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a84")="N" or session("a84")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Geração de Arquivo Unimed</title>
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

dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

sql="SELECT * FROM ttunimed "
%>
<%
caminho="c:\inetpub\wwwroot\rh\temp\"
nomefile="planilha_unimed_" & session.sessionid & ".vbs"
lote=caminho & nomefile
Set arquivo=CreateObject("Scripting.FileSystemObject")
Set leitura=arquivo.CreateTextFile(lote, true)

leitura.writeline "Dim objXL"
leitura.writeline "Set objXL = WScript.CreateObject(""Excel.Application"")"
leitura.writeline "objXL.Visible = TRUE"
leitura.writeline "objXL.WorkBooks.Add"
'leitura.writeline "objXL.Cells(1, 1).Value = ""Planilha para movimentação de Empresas - PMG"""
'leitura.writeline "objXL.Cells(3, 1).Value = ""Razão Social: """
'leitura.writeline "objXL.Cells(4, 1).Value = ""Data da Movimentação: """
'leitura.writeline "objXL.Cells(3, 4).Value = ""FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO - FIEO"""
'leitura.writeline "objXL.Cells(4, 4).Value = """ & formatdatetime(dtaccess(now()),2) & """"

leitura.writeline "objXL.Cells(1, 1).Value = ""Tipo Movto"""
leitura.writeline "objXL.Cells(1, 2).Value = ""Cód U' Seg"""
leitura.writeline "objXL.Cells(1, 3).Value = ""Código Empresa"""
leitura.writeline "objXL.Cells(1, 4).Value = ""Código Família"""
leitura.writeline "objXL.Cells(1, 5).Value = ""Relação Dep."""
leitura.writeline "objXL.Cells(1, 6).Value = ""Dígito"""
leitura.writeline "objXL.Cells(1, 7).Value = ""Data Nascimento"""
leitura.writeline "objXL.Cells(1, 8).Value = ""Sexo"""
leitura.writeline "objXL.Cells(1, 9).Value = ""Estado Civil"""
leitura.writeline "objXL.Cells(1, 10).Value = ""Data Inclusão/Exclusão"""
leitura.writeline "objXL.Cells(1, 11).Value = ""Plano/Prod."""
leitura.writeline "objXL.Cells(1, 12).Value = ""Nome Segurado"""
leitura.writeline "objXL.Cells(1, 13).Value = ""CPF Titular"""
leitura.writeline "objXL.Cells(1, 14).Value = ""Cidade"""
leitura.writeline "objXL.Cells(1, 15).Value = ""UF"""
leitura.writeline "objXL.Cells(1, 16).Value = ""Admissão"""
leitura.writeline "objXL.Cells(1, 17).Value = ""Nome da Mãe"""
leitura.writeline "objXL.Cells(1, 18).Value = ""Endereço"""
leitura.writeline "objXL.Cells(1, 19).Value = ""Número"""
leitura.writeline "objXL.Cells(1, 20).Value = ""Complemento"""
leitura.writeline "objXL.Cells(1, 21).Value = ""Bairro"""
leitura.writeline "objXL.Cells(1, 22).Value = ""CEP"""
leitura.writeline "objXL.Cells(1, 23).Value = ""PIS/PASEP"""
leitura.writeline "objXL.Cells(1, 24).Value = ""Matrícula"""
leitura.writeline "objXL.Cells(1, 25).Value = ""Lotação do Func."""
leitura.writeline "objXL.Cells(1, 26).Value = ""Declaração Nasc.Vivo"""
leitura.writeline "objXL.Cells(1, 27).Value = ""Cartão Nacional SUS"""
leitura.writeline "objXL.Cells(1, 28).Value = ""DDD Celular"""
leitura.writeline "objXL.Cells(1, 29).Value = ""Celular"""
leitura.writeline "objXL.Cells(1, 30).Value = ""Email"""

leitura.writeline "objXL.Range(""A1:AD1"").Select"
Bordas
leitura.writeline "objXL.Selection.Font.Bold = True"
leitura.writeline "objXL.Selection.Interior.ColorIndex = 15" 'gray
leitura.writeline "objXL.Selection.Interior.Pattern = 1 " 'xlSolid
leitura.writeline "objXL.Selection.Font.ColorIndex = 1" 'black

'leitura.writeline "objXL.Range(""A1:AD1"").Select"
'leitura.writeline "objXL.Selection.Font.Bold = True"
'CentralizaCelulas
leitura.writeline "objXL.Rows(""1:1"").RowHeight = 23"
leitura.writeline "objXL.Rows(""1:1"").Select"
leitura.writeline "objXL.Selection.HorizontalAlignment = -4108" 'xlCenter
leitura.writeline "objXL.Selection.VerticalAlignment = -4160" 'xlTop
leitura.writeline "objXL.Selection.WrapText = True"
leitura.writeline "objXL.Selection.Orientation = 0"
leitura.writeline "objXL.Selection.Font.Name = ""Arial"""
leitura.writeline "objXL.Selection.Font.Size = 8"
leitura.writeline "objXL.Columns(1).ColumnWidth = 6"
leitura.writeline "objXL.Columns(2).ColumnWidth = 7"
leitura.writeline "objXL.Columns(3).ColumnWidth = 10"
leitura.writeline "objXL.Columns(4).ColumnWidth = 10"
leitura.writeline "objXL.Columns(5).ColumnWidth = 35"
leitura.writeline "objXL.Columns(6).ColumnWidth = 6"
leitura.writeline "objXL.Columns(7).ColumnWidth = 8"
leitura.writeline "objXL.Columns(8).ColumnWidth = 16"
leitura.writeline "objXL.Columns(9).ColumnWidth = 16"
leitura.writeline "objXL.Columns(10).ColumnWidth = 10"
leitura.writeline "objXL.Columns(11).ColumnWidth = 12"
leitura.writeline "objXL.Columns(12).ColumnWidth = 12"
leitura.writeline "objXL.Columns(13).ColumnWidth = 12"
leitura.writeline "objXL.Columns(14).ColumnWidth = 12"
leitura.writeline "objXL.Columns(15).ColumnWidth = 12"
leitura.writeline "objXL.Columns(16).ColumnWidth = 12"
leitura.writeline "objXL.Columns(17).ColumnWidth = 12"
leitura.writeline "objXL.Columns(18).ColumnWidth = 7"
leitura.writeline "objXL.Columns(19).ColumnWidth = 14"
leitura.writeline "objXL.Columns(20).ColumnWidth = 14"
leitura.writeline "objXL.Columns(21).ColumnWidth = 20"
leitura.writeline "objXL.Columns(22).ColumnWidth = 15"
leitura.writeline "objXL.Columns(23).ColumnWidth = 7"
leitura.writeline "objXL.Columns(24).ColumnWidth = 14"
leitura.writeline "objXL.Columns(25).ColumnWidth = 14"
leitura.writeline "objXL.Columns(26).ColumnWidth = 14"
leitura.writeline "objXL.Columns(27).ColumnWidth = 14"
leitura.writeline "objXL.Columns(28).ColumnWidth = 10"
leitura.writeline "objXL.Columns(29).ColumnWidth = 14"
leitura.writeline "objXL.Columns(30).ColumnWidth = 24"

'leitura.writeline "objXL.Range(""A3:C3"").Select"
'CentralizaCelulas
'Esquerda	
'leitura.writeline "objXL.Range(""A4:C4"").Select"
'CentralizaCelulas
'Esquerda	
linha=2
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	rs.movefirst
	do while not rs.eof 
		leitura.writeline "objXL.Cells(" & linha & ", 1).Value = """ & rs("Tipo Movto") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 1).HorizontalAlignment = -4108" 'xlCenter
		leitura.writeline "objXL.Cells(" & linha & ", 2).Value = ""'" & rs("Cod U Seg") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 3).Value = ""'" & rs("Código Empresa") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 4).Value = ""'" & rs("Código Família") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 5).Value = ""'" & rs("Relação Dep") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 6).Value = """ & rs("Dígito") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 7).Value = """ & dtaccess(rs("Data Nascimento")) & """"
		leitura.writeline "objXL.Cells(" & linha & ", 8).Value = """ & rs("Sexo") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 9).Value = """ & rs("Estado Civil") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 10).Value = """ & dtaccess(rs("Data Incl/Excl")) & """"
		leitura.writeline "objXL.Cells(" & linha & ", 11).Value = """ & rs("Plano/Prod") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 12).Value = """ & rs("Nome Segurado") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 13).Value = ""'" & rs("CPF Titular") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 14).Value = """ & rs("cidade") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 15).Value = """ & rs("UF") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 16).Value = """ & dtaccess(rs("admissão")) & """"
		leitura.writeline "objXL.Cells(" & linha & ", 17).Value = """ & rs("Nome da Mãe") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 18).Value = ""'" & rs("Endereço") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 19).Value = """ & rs("Número") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 20).Value = """ & rs("Complemento") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 21).Value = """ & rs("Bairro") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 22).Value = ""'" & rs("CEP") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 23).Value = ""'" & rs("PIS/PASEP") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 24).Value = ""'" & rs("Matrícula") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 25).Value = """ & rs("Lotação do Func") & """"

		leitura.writeline "objXL.Cells(" & linha & ", 27).Value = """ & rs("Cartao Sus") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 28).Value = """ & rs("DDD Celular") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 29).Value = """ & rs("Celular") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 30).Value = """ & rs("Email") & """"
	linha=linha+1
	rs.movenext:	loop
	else 'rs.recordcount
	end if 'rs.recordcount
	rs.close
	leitura.writeline "objXL.Range(""A2:AD" & linha-1 & """).Select"
	Bordas
	leitura.writeline "objXL.Cells.Select"
	leitura.writeline "objXL.Selection.Font.Name = ""Tahoma"""
	leitura.writeline "objXL.Selection.Font.Size = 8"
	leitura.writeline "objXL.Cells.EntireColumn.AutoFit"
	leitura.writeline "objXL.Range(""A2"").Select"
	leitura.writeline "objXL.Columns(4).ColumnWidth = 10"
	leitura.close
	set leitura=nothing
	set arquivo=nothing

%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width='650'>
<tr>
	<td><p class=titulo>Movimentação de Inclusão/Exclusão/Alteração Unimed Seguros</td>
	<td><a href="../temp/<%=nomefile%>">Planilha Unimed</a></td>
</tr>
</table>
<%
rs.Open sql, ,adOpenStatic, adLockReadOnly
total=0
if rs.recordcount>0 then
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	if rs.fields(a)=">>>>>Pesquisar" then colorerr="<font color=red><b>" else colorerr="<font color=black>"
	response.write "<td class=""campor"" nowrap>" & colorerr & rs.fields(a) & "</td>"
	if rs.fields(a)=">>>>>Pesquisar" then mensagemfinal="Os cadastros desta planilha tem dados incorretos ou faltantes."
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
else 'rs.recordcount
	response.write "<b><font size=3>Não existe movimentação cadastrada.</b><br>"
end if 'rs.recordcount
rs.close
response.write "<p>"
termino=now()
duracao=(termino-inicio)
Response.write "<p>Inicio: " & inicio & "<br>Termino: " & termino & "<br>Duracao: " & formatdatetime(duracao,3)
response.write "<p><font size=3 color=red>" & mensagemfinal & "</font></p>"
%>

</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>