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
<title>Número/Evolução de Nº Funcionários Setor</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"><!--
function nome1() {	form.secao.value=form.nome.value; }
function secao1() {	form.nome.value=form.secao.value; }
--></script>
</head>
<body style="margin-left:20px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sessao=session.sessionid
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
anoinicial=2003
espacamento=5
if request.form="" then
sql="select e.codsecao, s.descricao from corporerm.dbo.psecao s, evolucao_setor e where e.codsecao=s.codigo collate database_default and e.ano>=" & anoinicial & " group by e.codsecao, s.descricao order by s.descricao "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="form" action="numerosetor.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleção de Setores para emissão do quadro comparativo</td>
</tr>
<tr>
	<td class=campo>Seção</td>
	<td class=campo><input type="text" name="secao" size="8" maxlength="8" onchange="secao1()"></td>
	<td class=campo>
		<select name="nome" class=a onchange="nome1()">
		<option value="0">Selecione um setor ou opção</option>
		<option value="todos">Soma de todos os Setores</option>
		<option value="pagina">Todos Setores agrupados por página</option>
<%
rs.movefirst
do while not rs.eof
%>
		<option value="<%=rs("codsecao")%>"> <%=rs("descricao")%></option>
<%
rs.movenext
loop
rs.close
%>
		</select>
	</td>
</tr>

<tr>
	<td class=campo>Entre os anos</td><td class=campo> de <input type="text" name="anoinicial" size="4" maxlength="4" value="<%=year(now)-2%>">
 	a <input type="text" name="anofinal" size="4" maxlength="4" value="<%=year(now)%>">                                 
	</td>
</tr>


<tr><td class=campo></td><td class=campo><input type="checkbox" name="imprimir" value="Apenas imprimir"> Apenas imprimir</td></tr>

<tr>
	<td class=campo colspan=3>&nbsp;
		<input type="submit" value="Visualizar" class=button name="B1">
	</td>
</tr>
</table>
</form>

<%
end if 'request.form=""

if request.form<>"" then
sql="select max(ano) as maxano from evolucao_setor"
sql="select max(ano) as maxano from evolucao_funcao"
rs.Open sql, ,adOpenStatic, adLockReadOnly
maxano=rs("maxano"):if isnull(maxano) or maxano="" then maxano=2003
rs.close
sql="select max(mes) as maxmes from evolucao_setor where ano=" & maxano
sql="select max(mes) as maxmes from evolucao_funcao where ano=" & maxano
rs.Open sql, ,adOpenStatic, adLockReadOnly
maxmes=rs("maxmes"):if isnull(maxmes) or maxmes="" then maxmes=1
rs.close
'if request.form("imprimir")="" then if request.form("secao")="pagina" then GeraDados maxano,maxmes
dataevo=dateserial(maxano,maxmes+1,1)-1
ano=year(dataevo):mes=month(dataevo)

if year(now)<>maxano or month(now)<>maxmes then
	novomes=maxmes+1
	novoano=maxano
	if novomes>12 then
		novomes=1
		novoano=novoano+1
	end if
	if request.form("imprimir")="" then if request.form("secao")="pagina" then GeraDados novoano,novomes
end if

'**********************************************
Sub GeraDados (var_ano,var_mes)
dataevo=dateserial(var_ano,var_mes+1,1)-1
ano=year(dataevo):mes=month(dataevo)
sql="delete from evolucao_setor where ano=" & ano & " and mes=" & mes:conexao.execute sql
sql="delete from evolucao_salario where ano=" & ano & " and mes=" & mes:conexao.execute sql
sql="delete from evolucao_situacao where ano=" & ano & " and mes=" & mes:conexao.execute sql
sql="delete from evolucao_funcao where ano=" & ano & " and mes=" & mes:conexao.execute sql

sql="select chapa from corporerm.dbo.pfunc where dataadmissao<='" & dtaccess(dataevo) & "' and (datademissao>='" & dtaccess(dataevo) & "' or datademissao is null) and (chapa<'10000' or chapa>='90000') and codsindicato<>'03' order by chapa"
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then

rs.movefirst
do while not rs.eof
	sql2="SELECT top 1 CHAPA, DTMUDANCA, Sum(SALARIO) AS salario FROM corporerm.dbo.PFHSTSAL " & _
	"GROUP BY CHAPA, DTMUDANCA HAVING CHAPA='" & rs("chapa") & "' AND DTMUDANCA<='" & dtaccess(dataevo) & "' ORDER BY DTMUDANCA DESC "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	sql3="insert into evolucao_salario (chapa, ano, mes, dtmudanca, salario) select '" & rs2("chapa") & "'," & ano & "," & mes & ",'" & dtaccess(rs2("dtmudanca")) & "'," & nraccess(rs2("salario")) & " "
	conexao.Execute Sql3, , adCmdText
	rs2.close
rs.movenext
loop

rs.movefirst
do while not rs.eof
	sql2="SELECT top 1 CHAPA, DTMUDANCA, CODSECAO FROM corporerm.dbo.PFHSTSEC " & _
	"WHERE CHAPA='" & rs("chapa") & "' AND DTMUDANCA<='" & dtaccess(dataevo) & "' ORDER BY CHAPA, DTMUDANCA DESC "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	sql3="insert into evolucao_setor (chapa, ano, mes, dtmudanca, codsecao) select '" & rs2("chapa") & "'," & ano & "," & mes & ",'" & dtaccess(rs2("dtmudanca")) & "','" & rs2("codsecao") & "' "
	conexao.Execute Sql3, , adCmdText
	rs2.close
rs.movenext
loop

rs.movefirst
do while not rs.eof
	sql2="SELECT top 1 CHAPA, DATAMUDANCA, NOVASITUACAO FROM corporerm.dbo.PFHSTSIT " & _
	"WHERE CHAPA='" & rs("chapa") & "' AND DATAMUDANCA<='" & dtaccess(dataevo) & "' ORDER BY CHAPA, DATAMUDANCA DESC "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
		sql3="insert into evolucao_situacao (chapa, ano, mes, datamudanca, novasituacao) select '" & rs2("chapa") & "'," & ano & "," & mes & ",'" & dtaccess(rs2("datamudanca")) & "','" & rs2("novasituacao") & "' "
		conexao.Execute Sql3, , adCmdText
	end if
	rs2.close
rs.movenext
loop

rs.movefirst
do while not rs.eof
	sql2="SELECT top 1 CHAPA, DTMUDANCA, CODFUNCAO FROM corporerm.dbo.PFHSTFCO " & _
	"WHERE CHAPA='" & rs("chapa") & "' AND DTMUDANCA<='" & dtaccess(dataevo) & "' ORDER BY CHAPA, DTMUDANCA DESC "
	'response.write sql2 & "<br>"
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	sql3="insert into evolucao_funcao (chapa, ano, mes, datamudanca, codfuncao) select '" & rs2("chapa") & "'," & ano & "," & mes & ",'" & dtaccess(rs2("dtmudanca")) & "','" & rs2("codfuncao") & "' "
	conexao.Execute Sql3, , adCmdText
	rs2.close
rs.movenext
loop

end if 'rs.recordcount>0
rs.close

sql3="SELECT f.CHAPA, Year(datademissao) AS ano, Month(datademissao) AS mes, f.CODSECAO, f.salario, f.codsituacao, f.DATADEMISSAO " & _
"FROM corporerm.dbo.PFUNC f WHERE (f.CHAPA<'10000' Or f.CHAPA>='90000') AND f.CODSITUACAO='D' AND f.CODTIPO='N' and f.codsindicato<>'03' " & _
"AND Year(datademissao)=" & ano & " AND Month(datademissao)=" & mes & " "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
	sql4="select chapa from evolucao_setor where chapa='" & rs("chapa") & "' and ano=" & rs("ano") & " and mes=" & rs("mes") & " and codsecao='" & rs("codsecao") & "' "
	rs2.Open sql4, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
	else
		sql2="INSERT INTO evolucao_setor ( CHAPA, ano, mes, CODSECAO, DTMUDANCA ) " & _
		"SELECT '" & rs("chapa") & "'," & rs("ano") & "," & rs("mes") & ",'" & rs("codsecao") & "','" & dtaccess(rs("datademissao")) & "'"
		conexao.Execute Sql2, , adCmdText
	end if
	rs2.close

	sql4="select chapa from evolucao_salario where chapa='" & rs("chapa") & "' and ano=" & rs("ano") & " and mes=" & rs("mes") & " "
	rs2.Open sql4, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
	else
		sql2="INSERT INTO evolucao_salario ( CHAPA, ano, mes, salario, DTMUDANCA ) " & _
		"SELECT '" & rs("chapa") & "'," & rs("ano") & "," & rs("mes") & "," & nraccess(rs("salario")) & ",'" & dtaccess(rs("datademissao")) & "'"
		conexao.Execute Sql2, , adCmdText
	end if
	rs2.close

	sql4="select chapa from evolucao_situacao where chapa='" & rs("chapa") & "' and ano=" & rs("ano") & " and mes=" & rs("mes") & " and novasituacao='" & rs("codsituacao") & "' "
	rs2.Open sql4, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
	else
		sql2="INSERT INTO evolucao_situacao ( CHAPA, ano, mes, novasituacao, DaTaMUDANCA ) " & _
		"SELECT '" & rs("chapa") & "'," & rs("ano") & "," & rs("mes") & ",'" & rs("codsituacao") & "','" & dtaccess(rs("datademissao")) & "'"
		conexao.Execute Sql2, , adCmdText
	end if
	rs2.close
rs.movenext
loop
end if
rs.close

End sub
'**********************************************

secao=request.form("secao")

anoinicial=request.form("anoinicial")
anofinal=request.form("anofinal")
todos=0
if secao="pagina" then
	sqlp="select s.codsecao from evolucao_setor s, corporerm.dbo.pfunc f where ano=" & maxano & " and mes=" & maxmes & " and s.chapa=f.chapa collate database_default and f.codsindicato<>'03' group by s.codsecao"
else
	sqlp="select s.codsecao from evolucao_setor s where s.codsecao='" & secao & "' group by s.codsecao"
end if
if secao="todos" then sqlp="select top 1 codsecao from evolucao_setor":todos=1

rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
rs3.movefirst
do while not rs3.eof
secao=rs3("codsecao")

sqls="select descricao from corporerm.dbo.psecao where codigo='" & secao & "' "
rs2.Open sqls, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 and todos=0 then nomesecao=rs2("descricao") else nomesecao="Todos"
rs2.close

sql1="SELECT f.CODSINDICATO, sal.ano, sal.mes, nativo=sum(case when novasituacao in ('A','F') and codtipo='N' then 1 else 0 end), " & _
"sativo=sum(case when novasituacao in ('A','F') and codtipo='N' then sal.salario else 0 end), nafast=sum(case when novasituacao not in ('A','F','D') and codtipo='N' then 1 else 0 end), " & _
"safast=sum(case when novasituacao not in ('A','F','D') and codtipo='N' then sal.salario else 0 end), nestag=sum(case when novasituacao in ('A','F') and codtipo='T' then 1 else 0 end), " & _
"sestag=sum(case when novasituacao in ('A','F') and codtipo='T' then sal.salario else 0 end), ndem=sum(case when novasituacao in ('D') and codtipo='N' then 1 else 0 end), " & _
"sdem=sum(case when novasituacao in ('D') and codtipo='N' then sal.salario else 0 end) " & _
"FROM ((corporerm.dbo.PFUNC f INNER JOIN evolucao_salario sal ON f.CHAPA collate database_default=sal.CHAPA) " & _
"INNER JOIN evolucao_setor sec ON (sal.mes=sec.mes) AND (sal.ano=sec.ano) AND (f.CHAPA collate database_default=sec.CHAPA))  " & _
"INNER JOIN evolucao_situacao sit ON (sal.mes=sit.mes) AND (sal.ano=sit.ano) AND (f.CHAPA collate database_default=sit.CHAPA) "
if todos=1 then sqla="WHERE (sal.ano>=" & anoinicial & " and sal.ano<=" & anofinal & ") and sec.codsecao not in ('03.1.999') " else sqla="WHERE sec.CODSECAO='" & secao & "' and (sal.ano>=" & anoinicial & " and sal.ano<=" & anofinal & ") "
sql2="GROUP BY f.CODSINDICATO, sal.ano, sal.mes " & _
"HAVING f.CODSINDICATO='01' " & _
"ORDER BY sal.ano, sal.mes "
sql=sql1 & sqla & sql2
rs.Open sql, ,adOpenStatic, adLockReadOnly

'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************
numero=0
rs.movefirst
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=500>
<th colspan=10>Evolução Funcional do setor: <%=nomesecao%></th>
<tr>
	<td class=titulop align="center" rowspan=2>Ano</td>
	<td class=titulop align="center" rowspan=2>Mês</td>
	<td class=titulop align="center" colspan=2 style="border-left:2 solid #000000">Ativos</td>
	<td class=titulop align="center" colspan=2 style="border-left:2 solid #000000">Estagiários</td>
	<td class=titulop align="center" style="border-left:2 solid #000000">Afastados</td>
	<td class=titulop align="center" style="border-left:2 solid #000000">Demitidos</td>
</tr>
<tr>
	<td class=titulop align="center" style="border-left:2 solid #000000">Nº</td>
	<td class=titulop align="center">R$ Base</td>
	<td class=titulop align="center" style="border-left:2 solid #000000">Nº</td>
	<td class=titulop align="center">R$ Base</td>
	<td class=titulop align="center" style="border-left:2 solid #000000">Nº</td>
	<td class=titulop align="center" style="border-left:2 solid #000000">Nº</td>
</tr>
<%
do while not rs.eof
redim preserve nano(numero):nano(numero)=rs("ano")
redim preserve nmes(numero):nmes(numero)=rs("mes")
redim preserve ativo(numero):ativo(numero)=rs("nativo")
if cdbl(rs("sativo"))<>0 then sativo=formatnumber(rs("sativo"),2) else sativo="-"
if cdbl(rs("sestag"))<>0 then sestag=formatnumber(rs("sestag"),2) else sestag="-"
if cdbl(rs("safast"))<>0 then safast=formatnumber(rs("safast"),2) else safast="-"
if cdbl(rs("sdem"))  <>0 then sdem  =formatnumber(rs("sdem"),2)   else sdem  ="-"

if rs("nativo")<>0 then nativo=formatnumber(rs("nativo"),0) else nativo="-"
if rs("nestag")<>0 then nestag=formatnumber(rs("nestag"),0) else nestag="-"
if rs("nafast")<>0 then nafast=formatnumber(rs("nafast"),0) else nafast="-"
if rs("ndem")  <>0 then ndem  =formatnumber(rs("ndem"),0)   else ndem  ="-"
%>
<tr>
	<td class="campop"><%=rs("ano")%></td>
	<td class="campop"><%=monthname(rs("mes"))%></td>
	<td class="campop" align="center" style="border-left:2 solid #000000"><%=nativo%></td>
	<td class="campop" align="right"><%=sativo%>&nbsp;</td>
	<td class="campop" align="center" style="border-left:2 solid #000000"><%=nestag%></td>
	<td class="campop" align="right"><%=sestag%>&nbsp;</td>
	<td class="campop" align="center" style="border-left:2 solid #000000"><%=nafast%></td>
	<td class="campop" align="center" style="border-left:2 solid #000000"><%=ndem%></td>
</tr>
<%
lano=rs("ano"):lmes=rs("mes")
rs.movenext
numero=numero+1
loop
rs.close
maximo=0
for ultimo=0 to ubound(ativo)
	if ativo(ultimo)>maximo then maximo=ativo(ultimo)
next
maximo=maximo+20
%>
</table>
<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td class=campo style="border-right: 1px solid #000000">
<%
texto="Nº de funcionários"
for a=1 to len(texto)
	response.write mid(texto,a,1) & "<br>"
next
%>
	</td>
	<%for volta=0 to ubound(nmes)%>
	<td class=campo>
	<%
	if ativo(volta)=0 then altura=0 else altura=int((ativo(volta)/maximo)*100)
	altura1=100-altura
	altura2=altura
	%>
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" height="250" width="100%">
	<tr><td height="<%=altura1%>%" align="center" valign=bottom class="campor"><%=ativo(volta)%></td>
	</tr>
	<tr><td class=fundo height="<%=altura2%>%" style="border: 1px solid #000000;font-size:1pt">&nbsp;</td>
	</tr>	
	</table>

	</td>
	<%next%>
</tr>

<tr>
	<td class=campo></td>
<%for volta=0 to ubound(nano)%>
	<td class=campo style="border-top: 1px solid #000000"><%=monthname(nmes(volta),1)%>
	<%if monthname(nmes(volta),1)="jan" then response.write "<br>/" & right(nano(volta),2)%>
	</td>
<%next%>
</tr>	

<tr>
	<td class=campo></td>
	<td class=campo colspan=24 align="center">Meses</td>
</tr>	
</table>
<%
response.write "<DIV style=""page-break-after:always""></DIV>"
if todos=0 then 
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse">
<th colspan=10>Funcionários alocados no setor: <%=nomesecao%></th>
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo width=150>Nome</td>
	<td class=titulo>Admissão</td>
	<td class=titulo>Tempo</td>
	<td class=titulo>Função</td>
	<td class=titulo>Tipo</td>
</tr>
<%
sql="select f.chapa, p.apelido, f.dataadmissao, f.datademissao, f.codtipo, c.nome as funcao, esi.novasituacao as codsituacao " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.pfuncao c, evolucao_situacao esi, evolucao_setor ese, evolucao_funcao efu " & _
"where f.codpessoa=p.codigo and efu.codfuncao=c.codigo collate database_default and esi.chapa=f.chapa collate database_default and esi.ano=" & lano & " and esi.mes=" & lmes & " " & _ 
"and ese.chapa=esi.chapa and ese.ano=esi.ano and ese.mes=esi.mes " & _
"and efu.chapa=esi.chapa and efu.ano=esi.ano and efu.mes=esi.mes " & _
"and esi.novasituacao<>'D' and f.codsindicato<>'03' " & _
"and ese.codsecao='" & secao & "' " & _
"order by p.apelido "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
tipofunc=""
if rs("codtipo")="T" then tipofunc="Estagiário"
if (rs("codsituacao")<>"A" and rs("codsituacao")<>"F" and rs("codsituacao")<>"Z") then tipofunc="Afast./Licenc."
if lano<>year(now) then
	d1=dateserial(lano,lmes+1,1)-1
	tempo=(d1-rs("dataadmissao"))/365.25
else
	d1=now()
	tempo=(d1-rs("dataadmissao"))/365.25
end if
tempo1=int(tempo)
tempo2=int(12*(tempo-tempo1))
if tempo1>0 then txttempo=tempo1 & " A. " else txttempo=""
if tempo2>0 then txttempo=txttempo & tempo2 & " M."
%>
<tr>
	<td class=campo><%=rs("chapa")%></td>
	<td class=campo><%=rs("apelido")%></td>
	<td class=campo><%=rs("dataadmissao")%></td>
	<td class=campo><%=txttempo%></td>
	<td class=campo><%=rs("funcao")%></td>
	<td class=campo><%=tipofunc%></td>
</tr>	
<%
rs.movenext
loop
end if 'rs.recordcount>0
rs.close
%>
</table>
<%
end if 'todos os setores

response.write "<DIV style=""page-break-after:always""></DIV>"
rs3.movenext
loop
rs3.close


end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
set rs2=nothing
set rs3=nothing
%>
</body>
</html>