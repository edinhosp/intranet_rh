<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a63")="N" or session("a63")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Geração de Arquivo SEFIP para autônomos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
inicio=now()
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	cmbmes=request.form("cmbmes")
    sql="DELETE FROM autonomo_sefip WHERE competencia='" & cmbmes & "' "
	conexao.execute sql
end if
%>
<p class=titulo>Geração de arquivo SEFIP de autônomos</p>
<%
if request.form="" then
mesfolha=month(dateserial(year(now),month(now)+1,1))
anofolha=year(dateserial(year(now),month(now)+1,1))
%>
<form method="POST" action="sefip.asp" name="form">
<p>Competência <select name="cmbmes">
<%
sqlmes="SELECT distinct cast(case when month(data_emissao)>9 then convert(char(2),month(data_emissao)) else '0'+cast(month(data_emissao) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_emissao])) AS Expr1 " & _
", Year([data_emissao]) , Month([data_emissao]) FROM autonomo_rpa " & _
"/*GROUP BY convert(char(2),case when month(data_emissao)<10 then '0' else '' end+Month([data_emissao]))+'/'+convert(char,Year([data_emissao])), Year([data_emissao]), Month([data_emissao]) */" & _
"ORDER BY Year([data_emissao]) DESC , Month([data_emissao]) DESC "
sqlmes="SELECT distinct cast(case when month(data_pagamento)>9 then convert(char(2),month(data_pagamento)) else '0'+cast(month(data_pagamento) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_pagamento])) AS Expr1 " & _
", Year([data_pagamento]) , Month([data_pagamento]) FROM autonomo_rpa " & _
"/*GROUP BY convert(char(2),case when month(data_pagamento)<10 then '0' else '' end+Month([data_pagamento]))+'/'+convert(char,Year([data_pagamento])), Year([data_pagamento]), Month([data_pagamento]) */" & _
"ORDER BY Year([data_pagamento]) DESC , Month([data_pagamento]) DESC "
response.write sqlmes
rsc.Open sqlmes, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
%>
	<option value="<%=rsc("expr1")%>"><%=rsc("expr1")%></option>
<%
rsc.movenext
loop
rsc.close
%>
	</select></p>
<p><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></p>
</form>
<% 
else

' ************* linha 00 ***************
c01 = "00" & Space(51) & "11"
c02 = Espaco2("73063166000120", 14)
c03 = Espaco2("FUNDACAO INSTITUTO DE ENSINO PARA OSASCO", 30)
c04 = Espaco2("LUIZ FERNANDO DA COSTA E SILVA", 20)
c05 = Espaco2("RUA NARCISO STURLINI 863", 50)
c06 = Espaco2("JD UMUARAMA", 20)
c07 = Espaco2("06018903", 8)
c08 = Espaco2("OSASCO", 20)
c09 = "SP001136519999"
c10 = Espaco2("rh@unifieo.br", 60)
Competencia = Right(cmbmes, 4) & Left(cmbmes, 2)
c11 = Competencia 
c12 = "905  " & Space(8) & "1" & Space(8) & Space(7) & "1" & "21867387000158" & Space(18)
c12 = "115" & "1" & "1" & Space(8) & "1" & Space(8) & space(7) & "1" & "21867387000158" & Space(18)
c13 = "*"
linha = c01 & c02 & c03 & c04 & c05 & c06 & c07 & c08 & c09 & c10 & c11 & c12 & c13
string_sql = "INSERT INTO autonomo_sefip (competencia, linha, linhasefip) " & _
"SELECT '" & cmbmes & "', '00', '" & linha & "'"
conexao.execute string_sql

' ************* linha 10 ***************
c01 = "10" & "1" & "73063166000120"
c02 = String(36, "0")
c03 = Espaco2("FUNDACAO INSTITUTO DE ENSINO PARA OSASCO", 40)
c04 = Espaco2("RUA NARCISO STURLINI 883", 50)
c05 = Espaco2("JD UMUARAMA", 20)
c06 = "06018903" & Espaco2("OSASCO", 20) & "SP" & "001136816000"
c07 = "N" & "8532500" & "N" & "00" & "0" & "1" & "639" & "0000" & "2305" & "10000"
c08 = String(15, "0") & String(15, "0") & String(15, "0") & "0"
c09 = String(14, "0") & Space(3) & Space(4) & Space(9) & String(45, "0") & Space(4)
c10 = "*"
linha = c01 & c02 & c03 & c04 & c05 & c06 & c07 & c08 & c09 & c10
string_sql = "INSERT INTO autonomo_sefip (competencia, linha, linhasefip) " & _
"SELECT '" & cmbmes & "', '10', '" & linha & "'"
conexao.execute string_sql

' ************* linha 90 ***************
linha = "90" & String(51, "9") & Space(306) & "*"
string_sql = "INSERT INTO autonomo_sefip (competencia, linha, linhasefip) " & _
"SELECT '" & cmbmes & "', '90', '" & linha & "'"
conexao.execute string_sql

' ************* linha 30 ***************
    linha1 = "30" & "1" & "73063166000120" & " " & Space(14)
    c06 = "" 'pis
    linha2 = Space(8) & "13"
    c09 = "" 'nome
    linha3 = Space(11) & Space(7) & Space(5) & Space(8) & Space(8)
    c15 = "0" & "0000" 'cbo
    c16 = String(15, "0") 'valor
    linha4 = String(15, "0") & "  " & "05"
    c20 = String(15, "0")
    linha5 = String(15, "0") & String(15, "0") & String(15, "0") & Space(98) & "*"

sql="SELECT cast(case when month(data_emissao)>9 then convert(char(2),month(data_emissao)) else '0'+cast(month(data_emissao) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_emissao])) AS competencia, " & _
"a.nome_autonomo, a.nit, '0'+Left([cbo],4) AS cbo2002, Sum([servico_prestado]+[outros_rendimentos]) AS valor_rem, Sum(r.desconto_inss) AS desconto_seg " & _
"FROM autonomo_rpa AS r INNER JOIN autonomo AS a ON r.id_autonomo = a.id_autonomo " & _
"GROUP BY cast(case when month(data_emissao)>9 then convert(char(2),month(data_emissao)) else '0'+cast(month(data_emissao) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_emissao])), a.nome_autonomo, a.nit, '0'+Left([cbo],4) " & _
"HAVING cast(case when month(data_emissao)>9 then convert(char(2),month(data_emissao)) else '0'+cast(month(data_emissao) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_emissao]))='" & cmbmes & "' "
sql="SELECT cast(case when month(data_pagamento)>9 then convert(char(2),month(data_pagamento)) else '0'+cast(month(data_pagamento) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_pagamento])) AS competencia, " & _
"a.nome_autonomo, a.nit, '0'+Left([cbo],4) AS cbo2002, Sum([servico_prestado]+[outros_rendimentos]) AS valor_rem, Sum(r.desconto_inss) AS desconto_seg " & _
"FROM autonomo_rpa AS r INNER JOIN autonomo AS a ON r.id_autonomo = a.id_autonomo " & _
"GROUP BY cast(case when month(data_pagamento)>9 then convert(char(2),month(data_pagamento)) else '0'+cast(month(data_pagamento) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_pagamento])), a.nome_autonomo, a.nit, '0'+Left([cbo],4) " & _
"HAVING cast(case when month(data_pagamento)>9 then convert(char(2),month(data_pagamento)) else '0'+cast(month(data_pagamento) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_pagamento]))='" & cmbmes & "' "
rsc.Open sql, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
	competencia=rsc("competencia")
	autonomo   =rsc("nome_autonomo")
	nit        =rsc("nit")
	cbo2002    =rsc("cbo2002")
	vrem       =formatnumber(rsc("valor_rem"),2)
    vrem       =replace(vrem,".","")
    vrem       =replace(vrem,",","")
    vrem       =numzero(vrem,15)
	vseg       =formatnumber(rsc("desconto_seg"),2)
    vseg       =replace(vseg,".","")
    vseg       =replace(vseg,",","")
    vseg       =numzero(vseg,15)
    string_sql = "INSERT INTO autonomo_sefip ( competencia, linha, linhasefip, pis ) " & _
	"SELECT '" & cmbmes & "', '30', '" & _
    linha1 & espaco2(textopuro(nit,2),11) & _
	linha2 & espaco2(textopuro(autonomo,2),70) & _
    linha3 & cbo2002 & vrem & _
    linha4 & vseg & _
    linha5 & "' " & ",'" & rsc("nit") & "' "
	conexao.execute string_sql
rsc.movenext
loop
rsc.close

sql="SELECT autonomo_sefip.linhasefip FROM autonomo_sefip " & _
"WHERE competencia='" & cmbmes & "' " & _
"ORDER BY competencia, linha, linhasefip"
sql="SELECT cast(case when month(data_emissao)>9 then convert(char(2),month(data_emissao)) else '0'+cast(month(data_emissao) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_emissao])) AS competencia, a.nome_autonomo, a.nit, '0'+Left([cbo],4) AS cbo2002, Sum([servico_prestado]+[outros_rendimentos]) AS valor_rem, Sum(r.desconto_inss) AS desconto_seg " & _
"FROM autonomo_rpa AS r INNER JOIN autonomo AS a ON r.id_autonomo = a.id_autonomo " & _
"GROUP BY cast(case when month(data_emissao)>9 then convert(char(2),month(data_emissao)) else '0'+cast(month(data_emissao) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_emissao])), a.nome_autonomo, a.nit, '0'+Left([cbo],4) " & _
"HAVING cast(case when month(data_emissao)>9 then convert(char(2),month(data_emissao)) else '0'+cast(month(data_emissao) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_emissao]))='" & cmbmes & "' "
sql="SELECT cast(case when month(data_pagamento)>9 then convert(char(2),month(data_pagamento)) else '0'+cast(month(data_pagamento) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_pagamento])) AS competencia, a.nome_autonomo, a.nit, '0'+Left([cbo],4) AS cbo2002, Sum([servico_prestado]+[outros_rendimentos]) AS valor_rem, Sum(r.desconto_inss) AS desconto_seg " & _
"FROM autonomo_rpa AS r INNER JOIN autonomo AS a ON r.id_autonomo = a.id_autonomo " & _
"GROUP BY cast(case when month(data_pagamento)>9 then convert(char(2),month(data_pagamento)) else '0'+cast(month(data_pagamento) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_pagamento])), a.nome_autonomo, a.nit, '0'+Left([cbo],4) " & _
"HAVING cast(case when month(data_pagamento)>9 then convert(char(2),month(data_pagamento)) else '0'+cast(month(data_pagamento) as char(1)) end as char(2))+'/'+convert(char(4),Year([data_pagamento]))='" & cmbmes & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalf=0:totalg=0
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='1' style='border-collapse: collapse' >"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>&nbsp;" & rs.fields(a).name & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	if rs.fields(a).type=5 then 
		conteudo=formatnumber(rs.fields(a),2) 
		response.write "<td align=""right"" class=""campor"">&nbsp;" & conteudo & "</td>"
	else 
		conteudo=rs.fields(a)
		response.write "<td class=""campor"">&nbsp;" & conteudo & "</td>"
	end if
	'response.write "<td><font size='1'>&nbsp;" &rs.fields(a) & rs.fields(a).type & "</td>"
next
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"
%>

<%
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="sefip" & ".re"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql="SELECT autonomo_sefip.linhasefip FROM autonomo_sefip " & _
	"WHERE competencia='" & cmbmes & "' " & _
	"ORDER BY competencia, linha, pis "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		leitura.writeline rs("linhasefip")
	rs.movenext
	loop
	rs.close
	termino=now()
	duracao=(termino-inicio)
	Response.write "<p class=realce><font size=1> Inicio: " & inicio & " Termino: " & termino & " Duracao: " & formatdatetime(duracao,3) & "</font></p>"
	leitura.close
	set leitura=nothing
	set arquivo=nothing

%>
<a href="..\temp\<%=nomefile%>">Arquivo SEFIP <%=cmbmes%></a>
<%
end if 'request.form 
%> 

</body>

</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>