<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a69")="N" or session("a69")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Geração de Pedido de Cesta-Basica</title>
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
sessao=session.sessionid

if request.form("gerar")<>"" then
	conexao.execute "delete from ttcbasica where sessao='" & sessao & "' and opcao='C' "

	dataliberacao=request.form("dataliberacao")
	mespedido=month(dataliberacao)+1:if mespedido>12 then mespedido=1
	mespedido=numzero(mespedido,2)
	dataliberacao=year(dataliberacao) & numzero(month(dataliberacao),2) & numzero(day(dataliberacao),2)
	
	data1=year(now) & numzero(month(now),2) & numzero(day(now),2)
	data2=formatdatetime(now,2)
	hora1=numzero(hour(now),2) & "." &  numzero(minute(now),2) & "." &  numzero(second(now),2)

	valorcb=formatnumber(request.form("valorcb"),2)
	taxacb=formatnumber(request.form("taxacb"),2)

'header pedido	
sql="insert into ttcbasica (sessao,campo1,campo2,campo3,registro, opcao) " & _
"select '" & sessao & "', '01', null, null, 'LSUP5" & espaco2(session("usuarioname"),8) & space(11) & data1 & hora1 & space(6) & espaco2("LAYOUT-04/12/2006",57) & space(165) & "','C' "
conexao.execute sql

'registro header
sql="insert into ttcbasica (sessao,campo1,campo2,campo3,registro, opcao) " & _
"select '" & sessao & "', '02', null, null, 'TA020A" & "0175760020" & espaco2("FUND.INST.ENSINO OSASCO",24) & space(6) & data1 & dataliberacao & "C" & space(16) & mespedido & space(19) & "04" & "33" & space(48) & "SUP   " & "000001','C' "
conexao.execute sql

'registro unidade
sql="insert into ttcbasica (sessao,campo1,campo2,campo3,registro,opcao) " & _
"select '" & sessao & "', '03', '01', null, 'TA022" & espaco2("NARCISO",26) & "AV  " & espaco2("FRANZ VOEGELLI",30) & "000300" & "5.ANDAR-RH" & _
espaco2("OSASCO",25) & espaco2("VILA YARA",15) & "06020SP" & espaco2(request.form("responsavel"),20)& "190" & space(7) & "000002','C'"
conexao.execute sql

sql="insert into ttcbasica (sessao,campo1,campo2,campo3,registro, opcao) " & _
"select '" & sessao & "', '03', '03', null, 'TA022" & espaco2("VILA YARA",26) & "AV  " & espaco2("FRANZ VOEGELLI",30) & "000300" & "5.ANDAR-RH" & _
espaco2("OSASCO",25) & espaco2("VILA YARA",15) & "06020SP" & espaco2(request.form("responsavel"),20)& "190" & space(7) & "000003','C'"
conexao.execute sql

sql="insert into ttcbasica (sessao,campo1,campo2,campo3,registro, opcao) " & _
"select '" & sessao & "', '03', '04', null, 'TA022" & espaco2("JD.WILSON",26) & "AV  " & espaco2("FRANZ VOEGELLI",30) & "000300" & "5.ANDAR-RH" & _
espaco2("OSASCO",25) & espaco2("VILA YARA",15) & "06020SP" & espaco2(request.form("responsavel"),20)& "190" & space(7) & "000004','C'"
conexao.execute sql

'registro de funcionarios
valorbeneficio=valorcb
valor1=int(valorcb)
valor2=int((valorcb-valor1+0.005)*100)
valorcb=numzero(valor1,7) & numzero(valor2,2)
taxacb2=nraccess(taxacb)
sequencia=5
sql="select a.chapa, f.nome, s.descricao, p.dtnascimento, left(f.codsecao,2) as unidade, f.codsecao " & _
"from ttcbasica_sel a, corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.ppessoa p " & _
"where f.codpessoa=p.codigo and f.codsecao=s.codigo and a.chapa=f.chapa collate database_default and a.sessao='" & sessao & "' and opcao='C' " & _
"order by f.codsecao, f.chapa "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalfu=rs.recordcount
rs.movefirst
do while not rs.eof
	'response.write "" & rs("chapa")
	datanasc=numzero(day(rs("dtnascimento")),2) & numzero(month(rs("dtnascimento")),2) & year(rs("dtnascimento"))
	if rs("unidade")="01" then unidade="NARCISO"
	if rs("unidade")="03" then unidade="VILA YARA"
	if rs("unidade")="04" then unidade="JD.WILSON"
	
	sql2="insert into ttcbasica (sessao,campo1,campo2,campo3,registro, taxa,opcao) " & _
	"select '" & sessao & "', '04', '" & rs("unidade") & rs("codsecao") & "', '" & rs("chapa") & "', " & _
	"'TA023" & espaco2(rs("descricao"),26) & numzero(rs("chapa"),12) & datanasc & space(18) & espaco2(unidade,26) & _
	"00101" & valorcb & "AE" & espaco2(rs("nome"),30) & space(17) & numzero(sequencia,6) & "', " & taxacb2 & ",'C'"
	conexao.execute sql2
rs.movenext
sequencia=sequencia+1
total=total+cdbl(valorbeneficio)
taxa=taxa+cdbl(taxacb)
loop
rs.close

valort1=int(total)
valort2=int((total-valort1+0.005)*100)
valort=numzero(valort1,7) & numzero(valort2,2)
'valortt=numzero(replace(replace(formatnumber(total,2),".",""),",","."),14)

sql="insert into ttcbasica (sessao,campo1,campo2,campo3,registro, opcao) " & _
"select '" & sessao & "', '05', null, null, 'TA029" & numzero(totalfu,8) & numzero(valort,14) & space(131) & numzero(sequencia,6) & "','C' "
conexao.execute sql

sql="select count(sessao) as totalarq from ttcbasica where sessao='" & sessao & "' and campo1 not in ('01','02','05','06') and opcao='C'"
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalarq=rs("totalarq")
rs.close

sql="insert into ttcbasica (sessao,campo1,campo2,campo3,registro,opcao) " & _
"select '" & sessao & "', '06', null, null, 'LSUP9" & "00000002" & "00000002" & numzero(totalarq,8) & space(277) & "','C' "
conexao.execute sql
	
sql="select * from ttcbasica where sessao='" & sessao & "' and opcao='C' order by campo1, right(campo2,8)"
end if
%>

<p class=titulo>Geração de Pedido para Cesta Básica</p>
<%
if request.form("gerar")="" then
dataliberacao=dateserial(year(now),month(now)+1,1)-1
anofolha=year(dateserial(year(now),month(now)+1,1))
sessao=session.sessionid

if request("acao")="excluir" then
	chapa=Request.QueryString("chapa")
	sql="delete from ttcbasica_sel where chapa='" & chapa & "' and sessao='" & sessao & "' and opcao='C' "
	conexao.execute sql
	manutencaocb=1
end if

if request.form("incluir")<>"" then
	if request.form("novachapa")<>"" then
		sql2="insert into ttcbasica_sel (sessao,chapa,opcao) " & _
		"select '" & sessao & "', '" & request.form("novachapa") & "','C' "
		conexao.execute sql2
	end if
	manutencaocb=1
end if

if request.form("alterar")="" then
	sm=724
	sql="select valor from iParametros where parametro='cblimite'"
	rs.open sql, ,adOpenStatic, adLockReadOnly
	limite=cdbl(rs("valor"))
	rs.close
	
	if manutencaocb<>1 then
		sql="DELETE FROM ttcbasica_sel where sessao='" & sessao & "' and opcao='C' "
		conexao.execute sql
		sql="SELECT f.CHAPA, f.CODSITUACAO, f.CODSINDICATO, f.CODTIPO, saltotal=case when dataadmissao>'2005-10-31' then salario else " & _
		"(power(1.050000,convert(integer,convert(float,('2005-10-31'-dtbase))/1095))-1) *([SALARIO]+(case when [043]=0 then " & sm & "*0.2 else 0 end)+(case when [175]>0 then [175] else 0 end)) +[SALARIO] end, jornadamensal/60 as jornmes " & _
		"FROM corporerm.dbo.PFUNC AS f LEFT JOIN ttcbasica_sel2 AS a ON f.CHAPA collate database_default=a.CHAPA " & _
		"inner join corporerm.dbo.pfcompl c on c.chapa=f.chapa " & _
		"WHERE f.CODSITUACAO Not In ('I','D') AND f.CODSINDICATO<>'03' AND f.CODTIPO='N' and c.opcbasica='C' order by f.chapa, codsecao "
		'response.write sql
		rs.Open sql, ,adOpenStatic, adLockReadOnly
		totalde=rs.recordcount
		rs.movefirst
		do while not rs.eof
			'salcomp=rs("saltotal")
			salcomp=(cdbl(rs("saltotal"))/cint(rs("jornmes")))*220
			'response.write rs("chapa") & "=" &  rs("saltotal") & "-" & rs("jornmes") & ": " & salcomp & "<br>"
			if salcomp<=limite then
				'response.write "---->Inseriu " & rs("chapa") & "<br>"
				sql2="insert into ttcbasica_sel (sessao,chapa, salario, limite, opcao) " & _
				"select '" & sessao & "', '" & rs("chapa") & "', " & nraccess(rs("saltotal")) & "," & nraccess(salcomp) & ", 'C'"
				conexao.execute sql2
				'response.write "7"
			end if
		rs.movenext
		loop
		rs.close
	end if
end if

sql="select c.chapa, f.nome, s.descricao, f.dataadmissao, f.codsituacao, c.salario, c.limite from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, ttcbasica_sel c " & _
"where f.chapa collate database_default=c.chapa and f.codsecao=s.codigo and c.sessao='" & sessao & "' and c.opcao='C' order by f.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" action="cb_pedido.asp">
<table border="1" cellpadding="0" cellspacing="1" style="border-collapse: collapse">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Setor</td>
	<td class=titulo>Admissão</td>
	<td class=titulo>Situação</td>
	<td class=titulo>Sal.Base</td>
	<td class=titulo>Sal.Calc</td>
	<td class=titulo>&nbsp;</td>
</tr>
<%
rs.movefirst
vezes=1
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("descricao")%></td>
	<td class=campo><%=rs("dataadmissao")%></td>
	<td class=campo><%=rs("codsituacao")%></td>
	<td class=campo align="right"><%=rs("salario")%></td>
	<td class=campo align="right"><%=rs("limite")%></td>
	<td class=campo align="center">&nbsp;
		<a href="cb_pedido.asp?acao=excluir&chapa=<%=rs("chapa")%>">
		<img border="0" src="../images/Trash.gif"></a>
	</td>
</tr>
<%
vezes=vezes+1
rs.movenext
loop
session("vezescb")=vezes-1
%>
<tr><td class=grupo colspan=8><%=rs.recordcount%> funcionários</td></tr>
<tr><td><input type="text" class="form_input" name="novachapa" size="5"></td>
	<td colspan=7><input type="submit" value="Incluir" class="button" name="incluir"></td>
</tr>
</table>

<%=""%>

<p>	Data Liberação: <input type="text" size="9" name="dataliberacao" value="<%=dataliberacao%>"><br>
Valor do Benefício: <input type="text" size=5 name="valorcb" value="1,00">
Taxa por cartão:  <input type="text" size=5 name="taxacb" value="0,00"><br>
Responsável Pedido: <input type="text" size="20" name="responsavel" value="ROGERIO MATEUS"></p>
<p><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></p>
</form>

<%
rs.close

'sm=380
sql="SELECT f.CHAPA, f.nome, f.datademissao, f.CODSITUACAO, f.CODSINDICATO, f.CODTIPO, saltotal= " & _
"(power(1.050000,convert(integer,convert(float,('2005-10-31'-dtbase))/1095))-1) *([SALARIO]+(case when [043]=0 then " & sm & "*0.2 else 0 end)+(case when [175]>0 then [175] else 0 end)) +[SALARIO] " & _
"FROM corporerm.dbo.PFUNC AS f LEFT JOIN ttcbasica_sel2 AS a ON f.CHAPA collate database_default= a.CHAPA " & _
"inner join corporerm.dbo.pfcompl c on c.chapa=f.chapa " & _
"WHERE datademissao between getdate()-30 and getdate() and Not f.CODSINDICATO='03' AND f.CODTIPO='N' and c.opcbasica='C' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="0" cellspacing="1" style="border-collapse: collapse">
<tr><td class=grupo colspan=3>Funcionários demitidos nos últimos 30 dias</td></tr>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Data Demissão</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
	if cdbl(rs("saltotal"))<=limite then
%>
<tr>
	<td class=campo><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("datademissao")%></td>
</tr>
<%
	end if
rs.movenext
loop
end if
rs.close
%>

<% else %>
<%

'----------------------------------------
sql="select * from ttcbasica where sessao='" & sessao & "' and opcao='C' order by campo1, campo2, substring(registro,159,6) "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalf=0:totalg=0
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='1' style='border-collapse: collapse' >"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>&nbsp;" & rs.fields(a).name & "</td>"
next
response.write "<td class=titulor>Total</td>"
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	if rs.fields(a).type=5 then 
		if isnull(rs.fields(a)) then conteudo=rs.fields(a) else conteudo=formatnumber(rs.fields(a),2) 
		response.write "<td align=""right"" class=""campor"">&nbsp;" & conteudo & "&nbsp;</td>"
	else 
		conteudo=rs.fields(a)
		response.write "<td class=""campor"">&nbsp;" & conteudo & "</td>"
	end if
next
response.write "<td class=""campor"" align=""right"">" & len(rs.fields("registro")) & "</td>"
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"
response.write limitecmtc
%>
Total de funcionarios: <%=totalfu%>
<br>Total do Pedido: <%=formatnumber(total,2) %>
<br>Tota da Taxa: <%=formatnumber(taxa,2)%>
<br>Total Geral: <%=formatnumber(total+taxa,2)%>

<%
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="tae" & textopuro(data2,2) & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql="select * from ttcbasica where sessao='" & sessao & "' and opcao='C' order by campo1, campo2, substring(registro,159,6)"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		leitura.writeline rs("registro")
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
<a href="..\temp\<%=nomefile%>">Arquivo Cesta Básica</a>
<%
end if 'request.form 
set rs=nothing
conexao.close
set conexao=nothing

%> 

</body>
</html>