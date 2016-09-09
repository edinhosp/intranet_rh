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
<title>Geração de Arquivo Medial</title>
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

conexao.execute "delete from ttmedial"
	
sql="INSERT INTO ttmedial (codmedial, tipo_mov, num_registro, beneficiario, parentesco, sexo, nome_mae, data_nascimento, plano, data_exclusao, up) " & _
"SELECT '10970' AS CodMedial, adm.oper AS Tipo_Mov, ab.CHAPA, ad.dependente, ad.parentesco, case when sexo='F' then 'Feminino' else 'Masculino' end, " & _
"ad.mae, ad.nascimento, adm.plano, adm.fvigencia, '0001' AS UP " & _
"FROM assmed_beneficiario ab, assmed_dep ad, assmed_dep_mudanca adm " & _
"WHERE ab.chapa=ad.chapa and ad.id_dep=adm.id_dep AND adm.oper='E' AND adm.empresa='M'"
conexao.execute sql
	
sql="INSERT INTO ttmedial (codmedial, tipo_mov, num_registro, beneficiario, parentesco, sexo, nome_mae, data_nascimento, plano, data_evento, up, cep_atendimento) " & _
"SELECT '10970' AS CodMedial, adm.oper AS Tipo_Mov, ab.CHAPA, ad.dependente, ad.parentesco, " & _
"case when sexo='F' then 'Feminino' else 'Masculino' end, ad.mae, ad.nascimento, adm.plano, ad.dt_evento, " & _
"'0001' AS UP, '06020-190' AS Expr2 " & _
"FROM assmed_beneficiario ab, assmed_dep ad, assmed_dep_mudanca adm " & _
"WHERE ab.chapa=ad.chapa and ad.id_dep=adm.id_dep AND adm.oper in ('I','A') AND adm.empresa='M'"
conexao.execute sql

sql="INSERT INTO ttmedial ( codmedial, tipo_mov, num_registro, beneficiario, parentesco, sexo, estado_civil, cpf, nome_mae, " & _
"endereco, numero, bairro, cidade, estado, cep, data_nascimento, plano, data_admissao, tp_contratacao, cep_atendimento, up, " & _
"[local], Departamento, rg, data_demissao, complemento, pis ) " & _
"SELECT '10970' AS CodMedial, am.oper AS Tipo_Mov, ab.CHAPA, ab.NOME, 'Titular' AS Expr1, case when sexo='F' then 'Feminino' else 'Masculino' end, " & _
"case estadocivil when 'I' then 'Divorciad' when 'C' then 'Casad' when 'S' then 'Solteir' when 'V' then 'Viuv' when 'D' then 'Desquit' when 'O' then 'Casad' else estadocivil end + case sexo when 'F' then 'a' else 'o' end, " & _
"p.CPF, m.MAE, p.RUA, p.NUMERO, p.BAIRRO, p.CIDADE, p.ESTADO, p.CEP, p.DTNASCIMENTO, am.plano, f.DATAADMISSAO, 'CLT' AS Expr3, " & _
"case when Left([codsecao],2)='01' then '06018-903' else '06020-190' end, '0001' AS Expr5, s.DESCRICAO, s.DESCRICAO AS Departamento, " & _
"p.cartidentidade, f.datademissao, p.complemento, f.pispasep " & _
"FROM assmed_beneficiario ab, corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.PSECAO s, assmed_mudanca am, qry_mae m " & _
"WHERE f.codpessoa=p.codigo and ab.chapa=f.chapa collate database_default and f.codsecao=s.codigo and ab.chapa=am.chapa and ab.chapa=m.chapa collate database_default " & _
"AND am.oper in ('I','A') AND am.empresa='M'"
conexao.execute sql
	
sql="INSERT INTO ttmedial ( codmedial, tipo_mov, num_registro, beneficiario, parentesco, plano, data_admissao, tp_contratacao, data_exclusao, data_demissao ) " & _
"SELECT '10970' AS CodMedial, am.oper AS Tipo_Mov, ab.CHAPA, ab.NOME, 'Titular' AS Expr1, am.plano, f.DATAADMISSAO, 'CLT' AS Expr3, am.fvigencia, f.datademissao " & _
"FROM assmed_beneficiario ab, corporerm.dbo.pfunc f, assmed_mudanca am " & _
"WHERE ab.chapa=f.chapa collate database_default and ab.chapa=am.chapa " & _
"AND am.oper='E' and am.empresa='M'"
conexao.execute sql

sql="INSERT INTO ttmedial ( codmedial, tipo_mov, num_registro, beneficiario, parentesco, plano  ) " & _
"SELECT '10970' AS CodMedial, '2ª Via', ab.CHAPA, ab.NOME, 'Titular' AS Expr1, am.plano " & _
"FROM assmed_beneficiario ab, assmed_mudanca am " & _
"WHERE ab.chapa=am.chapa AND am.oper='2' and am.empresa='M'"
conexao.execute sql

sql="INSERT INTO ttmedial (codmedial, tipo_mov, num_registro, beneficiario, parentesco, plano) " & _
"SELECT '10970' AS CodMedial, '2ª Via', ab.CHAPA, ad.dependente, ad.parentesco, adm.plano " & _
"FROM assmed_beneficiario ab, assmed_dep ad, assmed_dep_mudanca adm " & _
"WHERE ab.chapa=ad.chapa and ad.id_dep=adm.id_dep AND adm.oper='2' AND adm.empresa='M'"
conexao.execute sql
	
sql="SELECT * FROM ttmedial ORDER BY tipo_mov, num_registro, parentesco DESC"
%>
<%
caminho="c:\inetpub\wwwroot\rh\temp\"
nomefile="planilha_medial_" & session.sessionid & ".vbs"
lote=caminho & nomefile
Set arquivo=CreateObject("Scripting.FileSystemObject")
Set leitura=arquivo.CreateTextFile(lote, true)

leitura.writeline "Dim objXL"
leitura.writeline "Set objXL = WScript.CreateObject(""Excel.Application"")"
leitura.writeline "objXL.Visible = TRUE"
leitura.writeline "objXL.WorkBooks.Add"
leitura.writeline "objXL.Cells(1, 1).Value = ""Planilha para movimentação de Empresas - PMG"""
leitura.writeline "objXL.Cells(3, 1).Value = ""Razão Social: """
leitura.writeline "objXL.Cells(4, 1).Value = ""Data da Movimentação: """
leitura.writeline "objXL.Cells(3, 4).Value = ""FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO - FIEO"""
leitura.writeline "objXL.Cells(4, 4).Value = """ & formatdatetime(dtaccess(now()),2) & """"

leitura.writeline "objXL.Cells(7, 1).Value = ""Código"""
leitura.writeline "objXL.Cells(7, 2).Value = ""Tipo de Mov"""
leitura.writeline "objXL.Cells(7, 3).Value = ""Número de Registro"""
leitura.writeline "objXL.Cells(7, 4).Value = ""Grau Parentesco"""
leitura.writeline "objXL.Cells(7, 5).Value = ""Nome Beneficiário"""
leitura.writeline "objXL.Cells(7, 6).Value = ""Sexo"""
leitura.writeline "objXL.Cells(7, 7).Value = ""Estado Civil"""
leitura.writeline "objXL.Cells(7, 8).Value = ""CPF"""
leitura.writeline "objXL.Cells(7, 9).Value = ""RG"""
leitura.writeline "objXL.Cells(7, 10).Value = ""Plano Padr. Conf."""
leitura.writeline "objXL.Cells(7, 11).Value = ""Data Nascimento"""
leitura.writeline "objXL.Cells(7, 12).Value = ""Data Admissão"""
leitura.writeline "objXL.Cells(7, 13).Value = ""Data última promoção"""
leitura.writeline "objXL.Cells(7, 14).Value = ""Data do Evento"""
leitura.writeline "objXL.Cells(7, 15).Value = ""Data de Demissão"""
leitura.writeline "objXL.Cells(7, 16).Value = ""Data de Exclusão"""
leitura.writeline "objXL.Cells(7, 17).Value = ""CEP Atendimento"""
leitura.writeline "objXL.Cells(7, 18).Value = ""UP"""
leitura.writeline "objXL.Cells(7, 19).Value = ""Local"""
leitura.writeline "objXL.Cells(7, 20).Value = ""Departamento"""
leitura.writeline "objXL.Cells(7, 21).Value = ""Endereço"""
leitura.writeline "objXL.Cells(7, 22).Value = ""Número"""
leitura.writeline "objXL.Cells(7, 23).Value = ""Complemento"""
leitura.writeline "objXL.Cells(7, 24).Value = ""Bairro"""
leitura.writeline "objXL.Cells(7, 25).Value = ""Cidade"""
leitura.writeline "objXL.Cells(7, 26).Value = ""Estado"""
leitura.writeline "objXL.Cells(7, 27).Value = ""CEP"""
leitura.writeline "objXL.Cells(7, 28).Value = ""Nome da Mãe"""
leitura.writeline "objXL.Cells(7, 29).Value = ""Tipo Contratação"""
leitura.writeline "objXL.Cells(7, 30).Value = ""PIS/PASEP"""

leitura.writeline "objXL.Range(""A7:AD7"").Select"
Bordas
leitura.writeline "objXL.Selection.Font.Bold = True"
leitura.writeline "objXL.Selection.Interior.ColorIndex = 15" 'gray
leitura.writeline "objXL.Selection.Interior.Pattern = 1 " 'xlSolid
leitura.writeline "objXL.Selection.Font.ColorIndex = 1" 'black

leitura.writeline "objXL.Range(""A1:AD1"").Select"
leitura.writeline "objXL.Selection.Font.Bold = True"
CentralizaCelulas
leitura.writeline "objXL.Rows(""7:7"").RowHeight = 23"
leitura.writeline "objXL.Rows(""7:7"").Select"
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
leitura.writeline "objXL.Columns(26).ColumnWidth = 7"
leitura.writeline "objXL.Columns(27).ColumnWidth = 14"
leitura.writeline "objXL.Columns(28).ColumnWidth = 14"
leitura.writeline "objXL.Columns(29).ColumnWidth = 14"
leitura.writeline "objXL.Columns(30).ColumnWidth = 10"

leitura.writeline "objXL.Range(""A3:C3"").Select"
CentralizaCelulas
Esquerda	
leitura.writeline "objXL.Range(""A4:C4"").Select"
CentralizaCelulas
Esquerda	
linha=8
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	rs.movefirst
	do while not rs.eof 
		leitura.writeline "objXL.Cells(" & linha & ", 1).Value = """ & rs("codmedial") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 2).Value = """ & rs("tipo_mov") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 2).HorizontalAlignment = -4108" 'xlCenter
		leitura.writeline "objXL.Cells(" & linha & ", 3).Value = ""'" & rs("num_registro") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 4).Value = """ & rs("parentesco") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 5).Value = """ & rs("beneficiario") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 6).Value = """ & rs("sexo") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 7).Value = """ & rs("estado_civil") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 8).Value = ""'" & rs("cpf") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 9).Value = ""'" & rs("rg") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 10).Value = """ & rs("plano") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 11).Value = """ & dtaccess(rs("data_nascimento")) & """"
		leitura.writeline "objXL.Cells(" & linha & ", 12).Value = """ & dtaccess(rs("data_admissao")) & """"
		leitura.writeline "objXL.Cells(" & linha & ", 13).Value = """ & dtaccess(rs("data_promocao")) & """"
		leitura.writeline "objXL.Cells(" & linha & ", 14).Value = """ & dtaccess(rs("data_evento")) & """"
		leitura.writeline "objXL.Cells(" & linha & ", 15).Value = """ & dtaccess(rs("data_demissao")) & """"
		leitura.writeline "objXL.Cells(" & linha & ", 16).Value = """ & dtaccess(rs("data_exclusao")) & """"
		leitura.writeline "objXL.Cells(" & linha & ", 17).Value = """ & rs("cep_atendimento") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 18).Value = ""'" & rs("up") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 19).Value = """ & rs("local") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 20).Value = """ & rs("departamento") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 21).Value = """ & rs("endereco") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 22).Value = """ & rs("numero") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 23).Value = """ & rs("complemento") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 24).Value = """ & rs("bairro") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 25).Value = """ & rs("cidade") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 26).Value = """ & rs("estado") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 27).Value = """ & rs("cep") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 28).Value = """ & rs("nome_mae") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 29).Value = """ & rs("tp_contratacao") & """"
		leitura.writeline "objXL.Cells(" & linha & ", 30).Value = """ & rs("pis") & """"
	linha=linha+1
	rs.movenext:	loop
	else 'rs.recordcount
	end if 'rs.recordcount
	rs.close
	leitura.writeline "objXL.Range(""A8:AD" & linha-1 & """).Select"
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
	<td><p class=titulo>Movimentação de Inclusão/Exclusão/Alteração Medial Saúde</td>
	<td><a href="../temp/<%=nomefile%>">Planilha Medial</a></td>
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
	response.write "<td class=""campor"" nowrap>" & rs.fields(a) & "</td>"
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
%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>