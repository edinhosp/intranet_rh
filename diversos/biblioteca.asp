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

if request.form<>"" then
	dtinicio=dtaccess(request.form("dt_inicio"))
	dtfim=dtaccess(request.form("dt_fim"))

sql="SELECT f.CHAPA as Chapa, f.NOME as Nome, " & _
"Função=case when codtipo='T' then 'Estagiário' when codsindicato='03' then 'Professor' else 'Administrativo' end, " & _
"p.rua, p.numero, p.complemento, p.cep AS CEP, p.BAIRRO AS Bairro, p.CIDADE AS Cidade, p.TELEFONE1 AS Telefone, " & _
"f.dtdesligamento AS Validade, s.descricao AS Departamento " & _
"FROM corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.psecao s " & _
"WHERE f.codpessoa=p.codigo and f.codsecao=s.codigo and (f.CHAPA<'10000' Or f.CHAPA>'90000') " & _
"AND (f.dataadmissao  Between '" & dtinicio & "' And '" & dtfim & "' " & _
"or f.dtdesligamento Between '" & dtinicio & "' And '" & dtfim & "') " & _
"ORDER BY f.nome "
end if

if request.form="" then 
dt_inicio=dateserial(year(now),month(now),1)
dt_fim=dateserial(year(now),month(now)+1,1)-1
%>
<p class=titulo>Arquivo de movimentação para Biblioteca
<form method="POST" action="biblioteca.asp" name="form">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td class=titulo>&nbsp;&nbsp;Data Inicial</td>
		<td class=titulo>&nbsp;&nbsp;Data Final</td>
	</tr>
	<tr>
		<td class=titulo>&nbsp;&nbsp;<input type="text" name="dt_inicio" size="10" value="<%=dt_inicio%>"></td>
		<td class=titulo>&nbsp;&nbsp;<input type="text" name="dt_fim"    size="10" value="<%=dt_fim%>"   ></td>
	</tr>
	<tr>
		<td class=titulo colspan="2" align="center"><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></td>
	</tr>
</table>
</form>
<%
else
	'******* geracao arquivo *************
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="alan_" & textopuro(dtfim,2) & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	inicio=now()
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	separador=";"
	separador="#"
	cab="Chapa" & separador
	cab=cab & "Via" & separador
	cab=cab & "Nome" & separador
	cab=cab & "Função" & separador
	cab=cab & "Endereço" & separador
	cab=cab & "CEP" & separador
	cab=cab & "Bairro" & separador
	cab=cab & "Cidade" & separador
	cab=cab & "Telefone" & separador
	cab=cab & "Validade" & separador
	cab=cab & "Departamento" & separador
	cab=cab &  "Status" 
	leitura.writeline cab
	do while not rs.eof 
		chapa   =rs("chapa")    & separador
		via     =digito(rs("chapa"))  & separador
		nome    =rs("nome")     & separador
		funcao  =rs("função")   & separador
		endereco=rs("rua") & " " & rs("numero")
		if isnull(rs("complemento")) then
			ende=""
		else
			ende=" - " & rs("complemento")
		end if
		endereco=endereco & ende & separador
		cep     =rs("cep")      & separador
		bairro  =rs("bairro")   & separador
		cidade  =rs("cidade")   & separador
		secao   =rs("departamento") & separador
		telefone=rs("telefone") & separador
		validade=rs("validade") & separador
		if isnull(rs("validade")) then
			status="Admitido"
		else
			status="Demitido"
		end if
		registro = chapa & via & nome & funcao & endereco & cep & bairro & cidade & telefone & validade & secao & status
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

end if 'request.form 
%> 
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>