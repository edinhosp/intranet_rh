<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a92")="N" or session("a92")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Criação do Orçamento</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, rs, rs2, o(12), d(12)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
%>
<p class=titulo>Rotina Anual para a criação do orçamento contábil - RH</p>
<form method="POST" action="orcamento_criar.asp" name="form">
<table border="0" bordercolor=black cellpadding="2" cellspacing="1" style="border-collapse: collapse" width=300>
<tr>
	<td class=titulo height=35>1. Criar tabelas bases</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R1">
</tr>
<tr>
	<td class=titulo height=35>2. Calcular</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R2">
</tr>
<tr>
	<td class=titulo height=35>3. Montar tabela</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R3">
</tr>
<tr>
	<td class=titulo height=35>4. Gerar arquivo</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R4">
</tr>
</table>

</form>
<hr>
<%
datager="1/1/2009"
dissidio1=(1+(5/100))*100
dissidio2=(1+(1/100))*100
anoorc=2009

if request.form("R1")<>"" then
	dtoper=formatdatetime(now,2)
	sql1="delete from orc_base "
	conexao.execute sql1
	sql1="INSERT INTO orc_base (chapa, codtipo, codsindicato, dataadmissao, codsecao, codsituacao, salario, ats, adn, horasmes, dtnascimento) " & _
	"SELECT chapa, codtipo, codsindicato, dataadmissao, codsecao, codsituacao, salario, ats, adn, horasmes, dtnascimento " & _
	"FROM pfunc_ats WHERE codsituacao<>'D' "
	conexao.execute sql1
	response.write "<br>Gerou tabela base de funcionários..."
	
	sql1="delete from orc_13 "
	conexao.execute sql1
	sql1="INSERT INTO orc_13 (chapa, parc1, parc2, admissao, avos) " & _
	"select chapa, parc1=case when codsindicato='03' then 11 else case when month(dtnascimento)-1=0 then 12 else month(dtnascimento)-1 end end, 12 AS parc2, dataadmissao, 0 " & _
	"FROM orc_base "
	conexao.execute sql1
	response.write "<br>Gerou tabela base de 13º Salário..."
	
	sql1="delete from orc_ferias "
	conexao.execute sql1
	sql1="INSERT INTO orc_ferias ( chapa, vencferias, codsind, codsit, saldo, iniprog, fimprog, diasferias, abono, diasabono, parc13 ) " & _
	"SELECT o.CHAPA, f.DTVENCFERIAS AS vencferias, f.CODSINDICATO AS codsind, f.CODSITUACAO AS codsit, f.SALDOFERIAS AS saldo, f.INICPROGFERIAS1 AS iniprog, f.FIMPROGFERIAS1 AS fimprog, f.NRODIASFERIAS AS diasferias, f.QUERABONO AS abono, f.NRODIASABONO AS diasabono, f.QUER1APARC13O AS parc13 " & _
	"FROM orc_base AS o INNER JOIN corporerm.dbo.pfunc AS f ON o.CHAPA=f.CHAPA collate database_default "
	conexao.execute sql1
	sql1="update orc_ferias set iniprog='07/1/2009', fimprog='07/30/2009', diasferias=30 " & _
	"where codsind='03' and iniprog is null and codsit not in ('I','L','P') "
	conexao.execute sql1
'convert(char,year(data_acerto))+'/'+convert(char,month(data_acerto))+'/01'
	sql1="update orc_ferias set iniprog=convert(char,year(vencferias))+'/'+convert(char,month(vencferias)+2)+'/01', " & _
	"fimprog=convert(char,year(vencferias))+'/'+convert(char,month(vencferias)+2)+'/30', diasferias=30 " & _
	"where codsind='01' and iniprog is null and codsit not in ('I','L','P') "
	sql1="update orc_ferias set iniprog=DATEADD(m,2,convert(char,year(vencferias))+'/'+convert(char,month(vencferias))+'/01'), " & _
	"fimprog=dateadd(m,2,DATEADD(d, 29, convert(char,year(vencferias))+'/'+convert(char,month(vencferias))+'/01')), diasferias=30 " & _
	"where codsind='01' and iniprog is null and codsit not in ('I','L','P') "
	conexao.execute sql1
	response.write "<br>Gerou tabela base de Férias..."

	sql1="delete from orc_basepn":conexao.execute sql1
	sql1="INSERT INTO orc_basepn ( CHAPA, CODCCUSTO, valorN ) SELECT n.CHAPA, gc.CODCCUSTO, " & _
	"valorn=sum((case when n.id_nomeacao=80 then 16.42 else case when n.id_nomeacao=12 then 45.47 else case when n.id_nomeacao=16 then 69.33 else 35.34 end end end)*(ch*fator)) " & _
	"FROM cnv_atividade AS c INNER JOIN (n_indicacoes AS n INNER JOIN g2cursoeve AS gc ON n.coddoc=gc.coddoc) ON (n.codeve=c.codevento) " & _
	"AND (c.id_nomeacao=n.id_nomeacao) WHERE n.coddoc Is Not Null AND '12/1/2008' Between mand_ini And mand_fim GROUP BY n.CHAPA, gc.CODCCUSTO "
	conexao.execute sql1
	response.write "<br>Gerou tabela de nomeações..."
	
	sql1="delete from orc_baseps":conexao.execute sql1
	sql1="INSERT INTO orc_baseps ( CHAPA, CCUSTO, valorS ) " & _
	"SELECT s.CHAPA, e.CCUSTO, Sum(s.VALOR) AS valorS " & _
	"FROM (corporerm.dbo.PFSALCMP AS s INNER JOIN corporerm.dbo.pevento AS e ON s.CODEVENTO=e.CODIGO) INNER JOIN corporerm.dbo.PFUNC f ON s.CHAPA=f.CHAPA " & _
	"WHERE f.CODSINDICATO='03' AND f.CODSItuacao<>'D' " & _
	"GROUP BY s.CHAPA, e.CCUSTO "
	conexao.execute sql1
	sql1="insert into orc_baseps (chapa, ccusto, valors) values ('01057','01.3.702',5533.2)"
	conexao.execute sql1
	sql1="UPDATE orc_baseps INNER JOIN corporerm.dbo.pfunc f ON orc_baseps.CHAPA=f.CHAPA collate database_default SET CCUSTO=pfunc.codsecao " & _
	"WHERE orc_baseps.CCUSTO Is Null " 
	sql1="UPDATE orc_baseps SET CCUSTO=f.codsecao " & _
	"FROM orc_baseps INNER JOIN corporerm.dbo.pfunc f ON orc_baseps.CHAPA=f.CHAPA collate database_default  WHERE orc_baseps.CCUSTO Is Null "
	conexao.execute sql1
	response.write "<br>Gerou tabela de salários cursos... "

	sql1="delete from orc_basepos":conexao.execute sql1
	sql1="INSERT INTO orc_basepos ( CHAPA, CCUSTO, ValorP ) " & _
	"SELECT ff.CHAPA, e.CCUSTO, Sum(VALOR)/12 AS valorP " & _
	"FROM (corporerm.dbo.pevento AS e INNER JOIN corporerm.dbo.pffinanc AS ff ON e.CODIGO=ff.CODEVENTO) INNER JOIN corporerm.dbo.PFUNC AS f ON ff.CHAPA=f.CHAPA " & _
	"WHERE (ff.DTPAGTO>='12/1/2007' AND (e.CODIGO In ('792','831','870') Or e.CODIGO Between '951' And '982') AND f.CODSITUACAO<>'D') " & _
	"GROUP BY ff.CHAPA, e.CCUSTO;"
	conexao.execute sql1

	response.write "<br>Gerou tabela de salários pós-graduação... "
	
end if

if request.form("R2")<>"" then
	dtoper=formatdatetime(now,2)
	sql1="delete from orc_tabela "
	conexao.execute sql1
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01 ) " & _
	"SELECT codsecao, e1=case when substring(codsecao,4,1)='1' then 1 else case when substring(codsecao,4,1)='2' then 2 else 3 end end, count(chapa) as e2 " & _
	"FROM orc_base WHERE CODSINDICATO In ('01','02','03') " & _
	"GROUP BY CODSECAO, case when substring(codsecao,4,1)='1' then 1 else case when substring(codsecao,4,1)='2' then 2 else 3 end end "
	conexao.execute sql1
	sql1="update orc_tabela set p02=p01, p03=p01, p04=p01, p05=p01, p06=p01, p07=p01, p08=p01, p09=p01, p10=p01, p11=p01, p12=p01 "
	conexao.execute sql1
	response.write "<br><b>Calculou numero de funcionários no ano (1)(2)(3)..."
	
	sql1="update orc_ferias set m1=month(iniprog), m2=month(fimprog) where iniprog is not null"
	conexao.execute sql1
	sql1="update orc_ferias set d1=case when m2=m1 then convert(float,fimprog-iniprog+1) else " & _
	"convert(float,dateadd(d,-1,dateadd(m,1,convert(char,year(iniprog))+'/'+convert(char,month(iniprog))+'/01'))-iniprog+1) end, " & _
	"d2=case when m2=m1 then 0 else convert(float,fimprog-dateadd(d,1,convert(char,year(fimprog))+'/'+convert(char,month(fimprog))+'/01')) end " & _
	"where iniprog is not null "
	conexao.execute sql1
	sql1="update orc_ferias set dtpagto=iniprog-2 where iniprog is not null "
	conexao.execute sql1
	response.write "<br>Calculou periodos de férias..."
	
	sql1="update orc_base set p01=salario+ats+adn" : conexao.execute sql1
	sql1="delete from orc_base where chapa in (select chapa from orc_basepn) ":conexao.execute sql1
	sql1="delete from orc_base where chapa in (select chapa from orc_baseps) ":conexao.execute sql1
	sql1="delete from orc_base where chapa in (select chapa from orc_basepos) ":conexao.execute sql1
	sql1="INSERT INTO orc_base ( CHAPA, CODSECAO, SALARIO, codtipo, codsindicato, codsituacao,ats,adn ) " & _
	"SELECT u.CHAPA, u.ccusto, Sum(u.valorP) AS valor, 'N','03','A',0,0 FROM (" & _
	"SELECT CHAPA, CODCCUSTO as ccusto, valorN as valorP FROM orc_basepn " & _
	"union all " & _
	"SELECT CHAPA, CCUSTO, valorS * 1.225 FROM orc_baseps " & _
	"union all " & _
	"SELECT chapa, ccusto, valorp from orc_basepos " & _
	") AS u GROUP BY u.CHAPA, u.ccusto "
	conexao.execute sql1
	sql1="update orc_base set p01=salario+ats+adn" : conexao.execute sql1
	
	response.write "<br>Calculou salários professores..."

	sql1="update orc_base set p02=p01, p03=ceiling(p01*" & dissidio1 & ")/100 " : conexao.execute sql1
	sql1="update orc_base set p04=p03, p05=p03, p06=p03, p07=p03, p08=ceiling(p03*" & dissidio2 & ")/100 " : conexao.execute sql1
	sql1="update orc_base set p09=p08, p10=p08, p11=p08, p12=p08 " : conexao.execute sql1
	response.write "<br>Calculou salários anuais..."
	
	sql1="update orc_13 set avos=12 where admissao<=dateadd(d,15,'" & dtaccess(datager) & "') "
	conexao.execute sql1
	sql1="update orc_13 set avos=12-month(admissao) where admissao>'" & dtaccess(datager) & "' "
	conexao.execute sql1
	response.write "<br>Calculou períodos de 13º..."

	for a=1 to 12
		mes=numzero(a,2)
		sql1="update orc_ferias set v1=p" & mes & "/30*d1 " & _
		"select from orc_ferias f, orc_base b where f.chapa=b.chapa and m1=" & a & " "
		sql1="UPDATE orc_ferias SET v1=ceiling(((p" & mes & "/30*[d1])/3)*400)/100 from orc_ferias INNER JOIN orc_base ON orc_ferias.CHAPA=orc_base.CHAPA  WHERE orc_ferias.m1=" & a & " "
		conexao.execute sql1
		sql1="UPDATE orc_ferias SET v2=ceiling(((p" & mes & "/30*[d2])/3)*400)/100 from orc_ferias INNER JOIN orc_base ON orc_ferias.CHAPA=orc_base.CHAPA  WHERE orc_ferias.m2=" & a & " "
		conexao.execute sql1

		sql1="UPDATE orc_13 SET v1=ceiling((((p" & mes & "/12)*[avos])/2)*100)/100 from orc_13 INNER JOIN orc_base ON orc_13.CHAPA=orc_base.CHAPA  WHERE parc1=" & a & " "
		conexao.execute sql1
		sql1="UPDATE orc_13 SET v2=ceiling(((p" & mes & "/12)*[avos])*100)/100-v1 from orc_13 INNER JOIN orc_base ON orc_13.CHAPA=orc_base.CHAPA  WHERE parc2=" & a & " "
		conexao.execute sql1
	next
	response.write "<br>Calculou valores de férias e 13º..."
	
	sql1="update orc_base set d01=30, d02=30, d03=30, d04=30, d05=30, d06=30, d07=30, d08=30, d09=30, d10=30, d11=30, d12=30"
	conexao.execute sql1
	
	for a=1 to 12
		mes=numzero(a,2)
		sql1="UPDATE orc_base SET d" & mes & "=[d" & mes  & "]-[d1] from orc_base INNER JOIN orc_ferias ON orc_base.CHAPA=orc_ferias.CHAPA  WHERE orc_ferias.m1=" & a & " and year(orc_ferias.dtpagto)=" & anoorc & " "
		conexao.execute sql1
		sql1="update orc_base set v" & mes & "=(p" & mes & "/30)*d" & mes & " ":conexao.execute sql1
		sql1="update orc_base set v" & mes & "=ceiling(v" & mes & "*100)/100 ":conexao.execute sql1
	next
	response.write "<br>Calculos valores de salários..."

	sql1="delete from orc_temp":conexao.execute sql1
	sql1="INSERT INTO orc_temp ( CHAPA, tsalario ) SELECT CHAPA, Sum(salario) AS tsalario FROM orc_base GROUP BY CHAPA":conexao.execute sql1
	sql1="delete from orc_temp where tsalario<=0":conexao.execute sql1
	sql1="UPDATE orc_base SET perc=ceiling([salario]/[tsalario]*10000)/10000 from orc_base b INNER JOIN orc_temp t ON b.CHAPA=t.CHAPA  ":conexao.execute sql1
	
end if

if request.form("R3")<>"" then
	dtoper=formatdatetime(now,2)
	
	sql1="delete from orc_tabela where tipo in (11,12,13)":conexao.execute sql1
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
	"SELECT b.codsecao, case when substring(codsecao,4,1)='1' then 11 else case when substring(codsecao,4,1)='2' then 12 else 13 end end, " & _
	"Sum(b.v01), Sum(b.v02), Sum(b.v03), Sum(b.v04), Sum(b.v05), Sum(b.v06), " & _
	"Sum(b.v07), Sum(b.v08), Sum(b.v09), Sum(b.v10), Sum(b.v11), Sum(b.v12) " & _
	"FROM orc_base AS b " & _
	"GROUP BY b.codsecao, case when substring(codsecao,4,1)='1' then 11 else case when substring(codsecao,4,1)='2' then 12 else 13 end end "
	conexao.execute sql1
	response.write "<br><b>Montou tabela de salários (11)(12)(13)..."
	
	sql1="delete from orc_tabela where tipo in (31,32,33)":conexao.execute sql1
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO ) SELECT codsecao, case when substring(codsecao,4,1)='1' then 31 else case when substring(codsecao,4,1)='2' then 32 else 33 end end " & _
	"FROM orc_13 dt INNER JOIN orc_base b ON dt.CHAPA=b.CHAPA " & _
	"GROUP BY b.codsecao, case when substring(codsecao,4,1)='1' then 31 else case when substring(codsecao,4,1)='2' then 32 else 33 end end "
	conexao.execute sql1
	for a=1 to 12
		mes=numzero(a,2)
		sql1="UPDATE orc_tabela SET P" & mes & " = t.p" & mes & "+(v1*perc) " & _
		"from (orc_13 AS dt INNER JOIN orc_base AS b ON dt.CHAPA=b.CHAPA) INNER JOIN orc_tabela AS t ON b.codsecao=t.CODSECAO " & _
		"WHERE dt.parc1=" & a & " AND t.TIPO=case when substring(b.codsecao,4,1)='1' then 31 else case when substring(b.codsecao,4,1)='2' then 32 else 33 end end "
		conexao.execute sql1
		sql1="UPDATE orc_tabela SET P" & mes & " = t.p" & mes & "+(v2*perc) " & _
		"FROM (orc_13 AS dt INNER JOIN orc_base AS b ON dt.CHAPA=b.CHAPA) INNER JOIN orc_tabela AS t ON b.codsecao=t.CODSECAO " & _
		"WHERE dt.parc2=" & a & " AND t.TIPO=case when substring(b.codsecao,4,1)='1' then 31 else case when substring(b.codsecao,4,1)='2' then 32 else 33 end end "
		conexao.execute sql1
	next
	response.write "<br><b>Montou tabela de 13º salário (31)(32)(33)..."

	sql1="delete from orc_tabela where tipo in (21,22,23)":conexao.execute sql1
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO ) SELECT b.codsecao, case when substring(b.codsecao,4,1)='1' then 21 else case when substring(b.codsecao,4,1)='2' then 22 else 23 end end " & _
	"FROM orc_ferias AS f INNER JOIN orc_base AS b ON f.CHAPA = b.CHAPA " & _
	"GROUP BY b.codsecao, case when substring(b.codsecao,4,1)='1' then 21 else case when substring(b.codsecao,4,1)='2' then 22 else 23 end end "
	conexao.execute sql1
	for a=1 to 12
		mes=numzero(a,2)
		sql1="UPDATE orc_tabela SET P" &  mes & " = t.p" & mes & " + ( (v1+v2) * perc ) " & _
		"from orc_ferias AS f INNER JOIN (orc_base AS b INNER JOIN orc_tabela AS t ON b.codsecao=t.CODSECAO) ON f.CHAPA=b.CHAPA " & _
		"WHERE Month(dtpagto)=" & a & " AND t.TIPO=case when substring(b.codsecao,4,1)='1' then 21 else case when substring(b.codsecao,4,1)='2' then 22 else 23 end end and year(dtpagto)=" & anoorc & " "
		conexao.execute sql1
	next
	response.write "<br><b>Montou tabela de Férias (21)(22)(23)..."

	sql1="delete from orc_tabela where tipo between 41 and 49":conexao.execute sql1
	o(1)=11:o(2)=12:o(3)=13:o(4)=21:o(5)=22:o(6)=23:o(7)=31:o(8)=32:o(9)=33
	d(1)=41:d(2)=42:d(3)=43:d(4)=44:d(5)=45:d(6)=46:d(7)=47:d(8)=48:d(9)=49
	for a=1 to 9
		sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
		"SELECT CODSECAO, " & d(a) & " , ceiling([P01]*8)/100, ceiling([P02]*8)/100, ceiling([P03]*8)/100, ceiling([P04]*8)/100, ceiling([P05]*8)/100, " & _
		"ceiling([P06]*8)/100, ceiling([P07]*8)/100, ceiling([P08]*8)/100, ceiling([P09]*8)/100, ceiling([P10]*8)/100, ceiling([P11]*8)/100, " & _
		"ceiling([P12]*8)/100 FROM orc_tabela WHERE TIPO=" & o(a) & " "
		conexao.execute sql1
	next
	response.write "<br><b>Montou tabela de FGTS sobre Salários, Férias e 13º (41)-(49)..."

	sql1="delete from orc_tabela where tipo between 51 and 59":conexao.execute sql1
	d(1)=51:d(2)=52:d(3)=53:d(4)=54:d(5)=55:d(6)=56:d(7)=57:d(8)=58:d(9)=59
	for a=1 to 9
		sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
		"SELECT CODSECAO, " & d(a) & " , ceiling([P01]*1)/100, ceiling([P02]*1)/100, ceiling([P03]*1)/100, ceiling([P04]*1)/100, ceiling([P05]*1)/100, " & _
		"ceiling([P06]*1)/100, ceiling([P07]*1)/100, ceiling([P08]*1)/100, ceiling([P09]*1)/100, ceiling([P10]*1)/100, ceiling([P11]*1)/100, " & _
		"ceiling([P12]*1)/100 FROM orc_tabela WHERE TIPO=" & o(a) & " "
		conexao.execute sql1
	next
	response.write "<br><b>Montou tabela de PIS sobre Salários, Férias e 13º (51)-(59)..."

	sql1="delete from orc_tabela where tipo between 61 and 66":conexao.execute sql1
	o(1)=11:o(2)=12:o(3)=13
	d(1)=61:d(2)=62:d(3)=63
	for a=1 to 3
		sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
		"SELECT CODSECAO, " & d(a) & " , ceiling([P01]*11.11)/100, ceiling([P02]*11.11)/100, ceiling([P03]*11.11)/100, ceiling([P04]*11.11)/100, " & _
		"ceiling([P05]*11.11)/100, ceiling([P06]*11.11)/100, ceiling([P07]*11.11)/100, ceiling([P08]*11.11)/100, ceiling([P09]*11.11)/100, ceiling([P10]*11.11)/100, " & _
		"ceiling([P11]*11.11)/100, ceiling([P12]*11.11)/100 FROM orc_tabela WHERE TIPO=" & o(a) & " "
		conexao.execute sql1
	next
	response.write "<br><b>Montou tabela de Provisão de Férias (61)-(63)..."

	o(1)=61:o(2)=62:o(3)=63
	d(1)=64:d(2)=65:d(3)=66
	for a=1 to 3
		sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
		"SELECT CODSECAO, " & d(a) & " , ceiling([P01]*9)/100, ceiling([P02]*9)/100, ceiling([P03]*9)/100, ceiling([P04]*9)/100, ceiling([P05]*9)/100, " & _
		"ceiling([P06]*9)/100, ceiling([P07]*9)/100, ceiling([P08]*9)/100, ceiling([P09]*9)/100, ceiling([P10]*9)/100, ceiling([P11]*9)/100, " & _
		"ceiling([P12]*9)/100 FROM orc_tabela WHERE TIPO=" & o(a) & " "
		conexao.execute sql1
	next
	response.write "<br><b>Montou tabela de Encargos de Provisão de Férias (64)-(66)..."

	
	sql1="delete from orc_tabela where tipo between 71 and 76":conexao.execute sql1
	o(1)=11:o(2)=12:o(3)=13
	d(1)=71:d(2)=72:d(3)=73
	for a=1 to 3
		sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
		"SELECT CODSECAO, " & d(a) & " , ceiling([P01]*8.3333)/100, ceiling([P02]*8.3333)/100, ceiling([P03]*8.3333)/100, ceiling([P04]*8.3333)/100, " & _
		"ceiling([P05]*8.3333)/100, ceiling([P06]*8.3333)/100, ceiling([P07]*8.3333)/100, ceiling([P08]*8.3333)/100, ceiling([P09]*8.3333)/100, " & _
		"ceiling([P10]*8.3333)/100, ceiling([P11]*8.3333)/100, ceiling([P12]*8.3333)/100 FROM orc_tabela WHERE TIPO=" & o(a) & " "
		conexao.execute sql1
	next
	response.write "<br><b>Montou tabela de Provisão de 13º Salário (71)-(73)..."

	o(1)=71:o(2)=72:o(3)=73
	d(1)=74:d(2)=75:d(3)=76
	for a=1 to 3
		sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
		"SELECT CODSECAO, " & d(a) & " , ceiling([P01]*9)/100, ceiling([P02]*9)/100, ceiling([P03]*9)/100, ceiling([P04]*9)/100, ceiling([P05]*9)/100, " & _
		"ceiling([P06]*9)/100, ceiling([P07]*9)/100, ceiling([P08]*9)/100, ceiling([P09]*9)/100, ceiling([P10]*9)/100, ceiling([P11]*9)/100, " & _
		"ceiling([P12]*9)/100 FROM orc_tabela WHERE TIPO=" & o(a) & " "
		conexao.execute sql1
	next
	response.write "<br><b>Montou tabela de Encargos de Provisão de 13º Salário (74)-(76)..."

	sql1="delete from orc_vt":conexao.execute sql1
	sql1="delete from orc_tabela where tipo in (81,82,83)":conexao.execute sql1
	sql1="INSERT INTO orc_vt ( CHAPA, CODSECAO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12, valor_dia, v01, v02, v03, v04, v05, v06, v07, v08, v09, v10, v11, v12 ) " & _
	"SELECT fvt.CHAPA, b.CODSECAO, b.P01, b.P02, b.P03, b.P04, b.P05, b.P06, b.P07, b.P08, b.P09, b.P10, b.P11, b.P12, Sum([VALOR]*[NROVIAGENS]) AS valor_dia, 0,0,0,0,0,0,0,0,0,0,0,0 " & _
	"FROM ((corporerm.dbo.PVALETR AS vt INNER JOIN corporerm.dbo.PTARIFA AS t ON vt.CODTARIFA = t.CODIGO) INNER JOIN corporerm.dbo.PFVALETR AS fvt ON vt.CODIGO = fvt.CODLINHA) INNER JOIN orc_base AS b ON fvt.CHAPA collate database_default = b.CHAPA " & _
	"WHERE (((b.CODSITUACAO)<>'D') AND ((getdate()) Between [iniciovigencia] And [finalvigencia]) AND ((getdate()) Between [dtinicio] And [dtfim])) " & _
	"GROUP BY fvt.CHAPA, b.CODSECAO, b.P01, b.P02, b.P03, b.P04, b.P05, b.P06, b.P07, b.P08, b.P09, b.P10, b.P11, b.P12 "
	conexao.execute sql1
	d(1)=23:d(2)=19:d(3)=24:d(4)=22:d(5)=22:d(6)=23:d(7)=25:d(8)=23:d(9)=23:d(10)=23:d(11)=22:d(12)=24
	sql1="update orc_vt set "
	for a=1 to 12
		mes=numzero(a,2)
		sql1=sql1 & "v" & mes & "=valor_dia*" & d(a) & " "
		if a<12 then sql1=sql1 & ", "
	next
	conexao.execute sql1
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
	"SELECT CODSECAO, case when substring(codsecao,4,1)='1' then 81 else case when substring(codsecao,4,1)='2' then 82 else 83 end end, " & _
	"Sum(v01), Sum(v02), Sum(v03), Sum(v04), Sum(v05), Sum(v06), Sum(v07), Sum(v08), Sum(v09), Sum(v10), Sum(v11), Sum(v12) " & _
	"FROM orc_vt GROUP BY CODSECAO, case when substring(codsecao,4,1)='1' then 81 else case when substring(codsecao,4,1)='2' then 82 else 83 end end " 
	conexao.execute sql1
	response.write "<br><b>Montou tabela de Vale Transporte (81)-(83)..."
	
	sql1="delete from orc_tabela where tipo in (91,92,93)":conexao.execute sql1
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01 ) " & _
	"SELECT codsecao, case when substring(codsecao,4,1)='1' then 91 else case when substring(codsecao,4,1)='2' then 92 else 93 end end, Sum([perc]*2.29) " & _
	"FROM orc_base GROUP BY codsecao, case when substring(codsecao,4,1)='1' then 91 else case when substring(codsecao,4,1)='2' then 92 else 93 end end "
	conexao.execute sql1
	aum1=(1+(0/100))*100
	aum2=(1+(0/100))*100
	sql1="update orc_tabela set p02=p01, p03=ceiling(p01*" & aum1 & ")/100 where tipo in (91,92,93) " : conexao.execute sql1
	sql1="update orc_tabela set p04=p03, p05=p03, p06=p03, p07=p03, p08=ceiling(p03*" & aum2 & ")/100 where tipo in (91,92,93) " : conexao.execute sql1
	sql1="update orc_tabela set p09=p08, p10=p08, p11=p08, p12=p08 where tipo in (91,92,93) " : conexao.execute sql1
	response.write "<br><b>Montou tabela de PCMSO (91)-(93)..."
	
	sql1="delete from orc_tabela where tipo in (101,102,103)":conexao.execute sql1
	salmin=415:limite=salmin*5:limite2=limite*1.08
	cb=(55+2.13)*1.03:cb2=cb*1.03
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
	"SELECT orc_base.codsecao, case when substring(codsecao,4,1)='1' then 101 else case when substring(codsecao,4,1)='2' then 102 else 103 end end, " & _
	"Sum(case when [P01]<=" & nraccess(limite) & " then " & nraccess(cb) & " else 0 end), " & _
	"Sum(case when [P02]<=" & nraccess(limite) & " then " & nraccess(cb) & " else 0 end), " & _
	"Sum(case when [P03]<=" & nraccess(limite) & " then " & nraccess(cb) & " else 0 end), " & _
	"Sum(case when [P04]<=" & nraccess(limite) & " then " & nraccess(cb) & " else 0 end), " & _
	"Sum(case when [P05]<=" & nraccess(limite2) & " then " & nraccess(cb) & " else 0 end), " & _
	"Sum(case when [P06]<=" & nraccess(limite2) & " then " & nraccess(cb) & " else 0 end), " & _
	"Sum(case when [P07]<=" & nraccess(limite2) & " then " & nraccess(cb2) & " else 0 end), " & _
	"Sum(case when [P08]<=" & nraccess(limite2) & " then " & nraccess(cb2) & " else 0 end), " & _
	"Sum(case when [P09]<=" & nraccess(limite2) & " then " & nraccess(cb2) & " else 0 end), " & _
	"Sum(case when [P10]<=" & nraccess(limite2) & " then " & nraccess(cb2) & " else 0 end), " & _
	"Sum(case when [P11]<=" & nraccess(limite2) & " then " & nraccess(cb2) & " else 0 end), " & _
	"Sum(case when [P12]<=" & nraccess(limite2) & " then " & nraccess(cb2) & " else 0 end) " & _
	"FROM orc_base WHERE codtipo='N' AND codsindicato='01' " & _
	"GROUP BY codsecao, case when substring(codsecao,4,1)='1' then 101 else case when substring(codsecao,4,1)='2' then 102 else 103 end end "
	conexao.execute sql1
	response.write "<br><b>Montou tabela de Cesta Básica (101)-(103)..."

	vvacina=30
	sql1="delete from orc_tabela where tipo in (111,112,113)":conexao.execute sql1
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P03 ) " & _
	"SELECT codsecao, case when substring(codsecao,4,1)='1' then 111 else case when substring(codsecao,4,1)='2' then 112 else 113 end end, Sum(" & vvacina & ") " & _
	"FROM orc_base WHERE codsindicato<>'03' " & _
	"GROUP BY codsecao, case when substring(codsecao,4,1)='1' then 111 else case when substring(codsecao,4,1)='2' then 112 else 113 end end "
	conexao.execute sql1
	response.write "<br><b>Montou tabela de Vacinação (111)-(113)..."

	sql1="delete from orc_assmed":conexao.execute sql1
	sql1="delete from orc_tabela where tipo in (121,122,123,131,132,133)":conexao.execute sql1
	sql1="INSERT INTO orc_assmed ( chapa, Tipo, Ass, v00 ) " & _
	"SELECT m.chapa, 'Titular', case when empresa='O' then 'AO' else case when empresa='M' then 'AMM' else 'AMI' end end, Sum(p.valor) " & _
	"FROM assmed_planos p INNER JOIN assmed_mudanca m ON (p.plano=m.plano) AND (p.codigo=m.empresa) " & _
	"WHERE m.fvigencia>getdate() and empresa<>'D' GROUP BY m.chapa, case when empresa='O' then 'AO' else case when empresa='M' then 'AMM' else 'AMI' end end " & _
	"HAVING Sum(p.valor)>0 "
	conexao.execute sql1
	sql1="INSERT INTO orc_assmed ( chapa, Tipo, Ass, v00 ) " & _
	"SELECT d.chapa, 'Dependente', case when empresa='O' then 'AO' else case when empresa='M' then 'AMM' else 'AMI' end end, Sum(p.valor) " & _
	"FROM assmed_planos p INNER JOIN (assmed_dep d INNER JOIN assmed_dep_mudanca m ON d.id_dep=m.id_dep) ON (p.plano=m.plano) AND (p.codigo=m.empresa) " & _
	"WHERE m.fvigencia>getdate() and empresa<>'D' GROUP BY d.chapa, case when empresa='O' then 'AO' else case when empresa='M' then 'AMM' else 'AMI' end end " & _
	"HAVING Sum(p.valor)>0 "
	conexao.execute sql1
	
	m1=1:m2=m1*1.1
	i1=1.05:i2=i1*1.1
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12 ) " & _
	"SELECT b.codsecao, case when substring(codsecao,4,1)='1' then 121 else case when substring(codsecao,4,1)='2' then 122 else 123 end end, " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m1) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m1) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m1) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m1) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m2) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m2) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m2) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m2) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m2) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m2) & " else [v00]*" & nraccess(i1) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m2) & " else [v00]*" & nraccess(i2) & " end), " & _
	"Sum(case when [Ass]='AMM' then [v00]*" & nraccess(m2) & " else [v00]*" & nraccess(i2) & " end) " & _
	"FROM orc_assmed AS a INNER JOIN orc_base AS b ON a.chapa=b.CHAPA " & _
	"WHERE a.Ass in ('AMM','AMI') GROUP BY b.codsecao, case when substring(codsecao,4,1)='1' then 121 else case when substring(codsecao,4,1)='2' then 122 else 123 end end "
	conexao.execute sql1
	response.write "<br><b>Montou tabela de Assistência Médica (121)-(133)..."

	sql1="delete from orc_tabela where tipo in (141,142,151,152)":conexao.execute sql1
	valcbnatal=31.43:mescbnatal=12
	valunif=221.85:mesunif=1
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P12 ) " & _
	"SELECT CODSECAO, case when tipo=1 then 141 else case when tipo=2 then 142 else 143 end end, [P12]*" & nraccess(valcbnatal) & " FROM orc_tabela " & _
	"WHERE tipo in (1,2) "
	conexao.execute sql1
	response.write "<br><b>Montou tabela de Cesta de Natal (141)-(142)..."
	
	sql1="INSERT INTO orc_tabela ( CODSECAO, TIPO, P01 ) " & _
	"SELECT CODSECAO, case when tipo=1 then 151 else case when tipo=2 then 152 else 153 end end, [P01]*" & nraccess(valunif) & " FROM orc_tabela " & _
	"WHERE tipo in (1,2) "
	conexao.execute sql1
	response.write "<br><b>Montou tabela de Uniformes (151)-(152)..."
	
	for a=1 to 12
		mes=numzero(a,2)
		sql1="update orc_tabela set p" & mes & "=ceiling(p" & mes & "*100)/100 ":conexao.execute sql1
	next
end if

if request.form("R4")<>"" then
	dtoper=formatdatetime(now,2)
	inicio=now()
'**************************** arquivo saldus
	sql1="SELECT c.CODCONTA AS conta, Sum(ta.P01) AS S01, Sum(ta.P02) AS S02, Sum(ta.P03) AS S03, Sum(ta.P04) AS S04, Sum(ta.P05) AS S05, Sum(ta.P06) AS S06, Sum(ta.P07) AS S07, Sum(ta.P08) AS S08, Sum(ta.P09) AS S09, Sum(ta.P10) AS S10, Sum(ta.P11) AS S11, Sum(ta.P12) AS S12, substring(CODSECAO,2,1) AS filial, ta.CODSECAO " & _
	"FROM (orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO) INNER JOIN corporerm.dbo.CCONTA c ON ti.conta_saldus=c.REDUZIDO collate database_default " & _
	"WHERE ti.conta_saldus Is Not Null " & _
	"GROUP BY c.CODCONTA, substring(CODSECAO,2,1), ta.CODSECAO "
	sqlpre1="select top 15 * from (" & sql1 & ") t"
	
	componto=1
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="orc_saldus_" & anoorc & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql=sql1
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		conta=espaco2(rs("conta"),40)
		v01=rs("s01"):v02=rs("s02"):v03=rs("s03"):v04=rs("s04"):v05=rs("s05"):v06=rs("s06")
		v07=rs("s07"):v08=rs("s08"):v09=rs("s09"):v10=rs("s10"):v11=rs("s11"):v12=rs("s12")
		if componto=0 then
			v01=numzero(v01*100,20):v02=numzero(v02*100,20):v03=numzero(v03*100,20)
			v04=numzero(v04*100,20):v05=numzero(v05*100,20):v06=numzero(v06*100,20)
			v07=numzero(v07*100,20):v08=numzero(v08*100,20):v09=numzero(v09*100,20)
			v10=numzero(v10*100,20):v11=numzero(v11*100,20):v12=numzero(v12*100,20)
			z="                   0"
		else
			v01=espaco1(nraccess(v01),20):v02=espaco1(nraccess(v02),20):v03=espaco1(nraccess(v03),20)
			v04=espaco1(nraccess(v04),20):v05=espaco1(nraccess(v05),20):v06=espaco1(nraccess(v06),20)
			v07=espaco1(nraccess(v07),20):v08=espaco1(nraccess(v08),20):v09=espaco1(nraccess(v09),20)
			v10=espaco1(nraccess(v10),20):v11=espaco1(nraccess(v11),20):v12=espaco1(nraccess(v12),20)
			z="                   0"
		end if		
		registro=conta
		registro=registro & v01 & v02 & v03 & v04 & v05 & v06 & v07 & v08 & v09 & v10 & v11 & v12
		for a=1 to 12 : registro=registro & z : next
		registro=registro & space(9) & rs("filial") 'numzero(rs("filial"),10)
		registro=registro & espaco2(rs("codsecao"),25)
		registro=registro & space(25) 'departamento 25
		registro=registro & space(40) 'contagerencial 40
		registro=registro & anoorc
		leitura.writeline registro
	rs.movenext
	loop
	rs.close
	leitura.close
	set leitura=nothing
	set arquivo=nothing


'**************************** arquivo nucleus	
	sql1="SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '1/1/2009' AS data, Sum(ta.P01) AS valor FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P01)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '2/1/2009' AS data, Sum(ta.P02) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P02)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '3/1/2009' AS data, Sum(ta.P03) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P03)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '4/1/2009' AS data, Sum(ta.P04) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P04)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '5/1/2009' AS data, Sum(ta.P05) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P05)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '6/1/2009' AS data, Sum(ta.P06) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P06)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '7/1/2009' AS data, Sum(ta.P07) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P07)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '8/1/2009' AS data, Sum(ta.P08) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P08)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '9/1/2009' AS data, Sum(ta.P09) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P09)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '10/1/2009' AS data, Sum(ta.P10) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P10)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '11/1/2009' AS data, Sum(ta.P11) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P11)>0 " & _
"union all " & _
"SELECT ta.CODSECAO, ti.conta_nucleus AS conta, '12/1/2009' AS data, Sum(ta.P12) AS S01 FROM orc_tipo AS ti INNER JOIN orc_tabela AS ta ON ti.tipo = ta.TIPO GROUP BY ta.CODSECAO, ti.conta_nucleus HAVING ti.conta_nucleus Is Not Null AND Sum(ta.P12)>0 "
	sql2=sql1
	sql1="select * from (" & sql2 & ") t"
	sqlpre2="select top 15 * from (" & sql1 & ") t"
	
	componto=1
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile2="orc_nucleus_" & anoorc & ".txt"
	lote=caminho & nomefile2
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql=sql1
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		ccusto=espaco2(rs("codsecao"),25)
		cnatureza=espaco2(rs("conta"),10)
		dia=numzero(day(rs("data")),2)
		m=numzero(cint(month(rs("data"))),2)
		y=numzero(year(rs("data")),4)
		data=dia&"/"&m&"/"&y
		valor=rs("valor"):valort=valor
		centavo=int((valor-int(valor))*100+.05)
		unidade=numzero(int(valor),15)
		if componto=0 then
			valor="0" & unidade & centavo & "00"
		else
			valor=unidade & "." & centavo & "00"
		end if
		valor=espaco1(nraccess(valort),20)
		registro=ccusto & cnatureza & data & valor
		leitura.writeline registro
	rs.movenext
	loop
	rs.close
	leitura.close
	set leitura=nothing
	set arquivo=nothing
	
	termino=now()
	duracao=(termino-inicio)
	Response.write "<p class=realce><font size=1> Inicio: " & inicio & " Termino: " & termino & " Duracao: " & formatdatetime(duracao,3) & "</font></p>"
	
	%>
<br>
<a href="..\temp\<%=nomefile%>">Arquivo Orçamento Contábil de Folha (Saldus)</a>
<br>
<a href="..\temp\<%=nomefile2%>">Arquivo Orçamento RH Folha (Nucleus)</a>
	<%
end if

			if request.form("R4")<>"" then
'*************** inicio teste **********************
for b=1 to 2
	if b=1 then sql=sqlpre1 else sql=sqlpre2
	texto="Amostra do arquivo "
	if b=1 then texto=texto & "Saldus" else texto=texto & "Nucleus"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if request.form<>"" then
	response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
	response.write "<tr><td colspan=" & rs.fields.count & "><b>" & texto & "</td><tr>"
	response.write "<tr>"
	for a=0 to rs.fields.count-1
		response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
	next
	response.write "</tr>"
	if rs.recordcount>0 then rs.movefirst
	do while not rs.eof 
	response.write "<tr>"
	for a= 0 to rs.fields.count-1
		response.write "<td class=""campor"" nowrap>" & rs.fields(a) & "</td>"
	next
	response.write "</tr>"
	rs.movenext
	loop
	rs.close
	response.write "</table>"
	response.write "<p>"
	end if
next 
'*************** fim teste **********************

			end if
%>

</body>
</html>
<%

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>