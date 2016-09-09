<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Geração de Arquivo Biblioteca</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

sql="SELECT f.CHAPA as Chapa, f.NOME as Nome, " & _
"Função=case when codtipo='T' then 'Estagiário' when codsindicato='03' then 'Professor' else 'Administrativo' end, " & _
"p.rua, p.numero, p.complemento, p.cep AS CEP, p.BAIRRO AS Bairro, p.CIDADE AS Cidade, p.TELEFONE1 AS Telefone, " & _
"f.dtdesligamento AS Validade, s.descricao AS Departamento " & _
"FROM corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.psecao s " & _
"WHERE f.codpessoa=p.codigo and f.codsecao=s.codigo and (f.CHAPA<'10000' Or f.CHAPA>'90000') " & _
"AND (f.codsituacao<>'D') " & _
"ORDER BY f.nome "

dtinicio=dateserial(year(now),month(now),1)
dtfim=dateserial(year(now),month(now)+1,1)-1

	'******* geracao arquivo *************
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="alan_" & textopuro(dtfim,2) & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	inicio=now()
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	cab=chr(34) & "Chapa" & chr(34) & ";"
	cab=cab & chr(34) & "Nome" & chr(34) & ";"
	cab=cab & chr(34) & "Função" & chr(34) & ";"
	cab=cab & chr(34) & "Endereço" & chr(34) & ";"
	cab=cab & chr(34) & "CEP" & chr(34) & ";"
	cab=cab & chr(34) & "Bairro" & chr(34) & ";"
	cab=cab & chr(34) & "Cidade" & chr(34) & ";"
	cab=cab & chr(34) & "Telefone" & chr(34) & ";"
	cab=cab & chr(34) & "Validade" & chr(34) & ";"
	cab=cab & chr(34) & "Departamento" & chr(34) & ";"
	cab=cab & chr(34) & "Status" & chr(34)
	leitura.writeline cab
	do while not rs.eof 
		chapa   =chr(34) & rs("chapa")    & chr(34) & ";"
		nome    =chr(34) & rs("nome")     & chr(34) & ";"
		funcao  =chr(34) & rs("função")   & chr(34) & ";"
		endereco=chr(34) & rs("rua") & " " & rs("numero")
		if isnull(rs("complemento")) then
			ende=""
		else
			ende=" - " & rs("complemento")
		end if
		endereco=endereco & ende & chr(34) & ";"
		cep     =chr(34) & rs("cep")      & chr(34) & ";"
		bairro  =chr(34) & rs("bairro")   & chr(34) & ";"
		cidade  =chr(34) & rs("cidade")   & chr(34) & ";"
		telefone=chr(34) & rs("telefone") & chr(34) & ";"
		validade=chr(34) & rs("validade") & chr(34) & ";"
		secao   =chr(34) & rs("departamento") & chr(34) & ";"
		if isnull(rs("validade")) then
			status=chr(34) & "Admitido" & chr(34)
		else
			status=chr(34) & "Demitido" & chr(34)
		end if
		registro = chapa & nome & funcao & endereco & cep & bairro & cidade & telefone & validade & secao & status
		leitura.writeline registro
	rs.movenext
	loop
	rs.close
	termino=now()
	duracao=(termino-inicio)
	'Response.write "Inicio: " & inicio & "<br>Termino: " & termino & "<br>Duracao: " & formatdatetime(duracao,3)
	leitura.close
	set leitura=nothing
	set arquivo=nothing

%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width='650'>
	<tr>
		<td><p class=titulo>Arquivo de Movimentação para Biblioteca</td>
		<td><a href="../temp/<%=nomefile%>">Arquivo</a></td>
	</tr>
</table>
<%
rs.Open sql, ,adOpenStatic, adLockReadOnly
total=0
rs.movefirst
response.write "<table border='1' cellpadding='0' cellspacing='1' style='border-collapse:collapse' width='650'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class="campor">" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"
%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>