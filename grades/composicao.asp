<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a82")="N" or session("a82")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Composição de Salário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, chapach, rs, rs1
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
if request.form<>"" then
	if request.form("B3")<>"" then
		finaliza=1
	else
		finaliza=0
	end if
end if
if finaliza=0 then
%>
<p class=titulo>Seleção para impressão de composição salarial</p>
<form method="POST" action="composicao.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=550>
<tr>
	<td class=titulo>Campus/Pessoa</td>
	<td class=titulo>Database Tabela</td>
	<td class=titulo>Database Grade</td>
</tr>
<tr>
	<td class=titulo>
	<select size="1" name="D1">
		<option value="Todos" <%if request.form("D1")="Todos" then response.write "selected"%> >Todos</option>
		<option value="Narciso" <%if request.form("D1")="Narciso" then response.write "selected"%> >Todos-Campus Narciso</option>
		<option value="Yara" <%if request.form("D1")="Yara" then response.write "selected"%> >Todos-Campus V.Yara</option>
<%
sql1="select chapa, nome from dc_professor where (codsituacao<>'D' or chapa='00753') order by nome "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
if request.form("D1")=rs("chapa") then temp="selected" else temp=""
%>
<option value="<%=rs("chapa")%>" <%=temp%> ><%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
if request.form("databasegrade")="" then
	sql1="select max(data) as data from csd_obs where data<=getdate() "
	databasegrade=formatdatetime(now,2)
else
	sql1="select max(data) as data from csd_obs where data<='" & dtaccess(request.form("databasegrade")) & "' "
	databasegrade=request.form("databasegrade")
end if
rs.Open sql1, ,adOpenStatic, adLockReadOnly
datatabela=rs("data")
rs.close
%>			
	</select>
	</td>
	<td class=titulo><input type="text" name="databasetabela" size="12" value="<%=datatabela%>"></td>
	<td class=titulo><input type="text" name="databasegrade" size="12" value="<%=databasegrade%>" onchange="javascript:submit();"></td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=550>
<tr><td align="center" class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3">
</td></tr>
</table>
<input type="text" class="a" name="ultimo"  value="<%=session("ultimacomposicao")%>" size="6" > 
</form>
<hr>
<%
end if 'finaliza=0

if finaliza=1 then
temp=" and dc.chapa in (select chapa collate database_default from quem_nomeacoes where tipo='RHT') "
temp=" and dc.chapa in (select distinct chapa1 collate database_default from g2ch where coddoc='ARQ' and perlet='2016/1' ) "
temp=" and dc.chapa in (select chapa collate database_default from corporerm.dbo.pfunc where dataadmissao>'20160801' and codtipo='N' and codsindicato='03') "
temp=" and dc.chapa>'" & request.form("ultimo") & "' " 'temp=" and dc.chapa>'00796' "
temp=" and dc.chapa in (select chapa1 collate database_default from achapa) "

	dtsalario=request.form("databasetabela")
	dtgrade=request.form("databasegrade")
	chapa=request.form("d1")
	if chapa="Todos" then
		sql2=" " & temp
	elseif chapa="Narciso" then
		sql2=" and left(codsecao,2)='01' " & temp
	elseif chapa="Yara" then
		sql2=" and left(codsecao,2)='03' " & temp
	else
		sql2=" and chapa='" & chapa & "' "
	end if
	
	sql1="SELECT top 50 dc.CHAPA, dc.NOME, dc.DATAADMISSAO, dc.CODFUNCAO, dc.FUNCAO, dc.CODSECAO, s.descricao as SECAO, codsituacao, " & _
"titulacao=case when ge='Sim' then 'GE' else titulacaopaga end, dc.CODNIVELSAL, dc.GE, dc.GRAUINSTRUCAO, dc.titulacaopaga, dc.INSTRUCAOmec, inicioprofessor " & _
"FROM dc_professor dc, corporerm.dbo.psecao s " & _
"WHERE dc.codsecao=s.codigo and (dc.CODSITUACAO<>'D' or chapa='99999') and dc.codtipo in ('N') " & sql2 & " order by dc.chapa  "

	'response.write sql1
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof
	titulacao=rs("titulacao")
	nivel=rs("codnivelsal")
	ttaulas=0
	ttjornada=0
	ttsalario=0
	anoc=year((dtgrade))
	if month((dtgrade))<8 then pl=1 else pl=2
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=650>
<tr><td class=titulo align="center" colspan=6>Montagem de Salário Camposto - <%=anoc&"/"&pl%></td><td class=fundo><%=rs.absoluteposition%></td></tr>
<tr><td class=titulor align="center">Chapa</td>
	<td class=titulor align="center">Nome</td>
	<td class=titulor align="center">Função</td>
	<td class=titulor align="center">Seção</td>
	<td class=titulor align="center">Titulação</td>
	<td class=titulor align="center">Nivel</td>
	<td class=titulor align="center">Admissao</td>
</tr>
<tr><td class="campoa"><b><%=rs("chapa")%></b></td>
	<td class="campoa"><b><%=rs("nome")%></b></td>
	<td class="campoa"r><%=rs("funcao")%>&nbsp;</td>
	<td class="campoa"r><%=rs("codsecao") & "-" & rs("secao")%>&nbsp;</td>
	<td class="campoa"r><%=rs("titulacaopaga") & "-" & rs("instrucaomec")%></td>
	<td class="campoa"r><%=rs("codnivelsal")%>&nbsp;<%=titulacao%></td>
	<td class="campoa"r><%=rs("dataadmissao")%> (<%=rs("codsituacao")%>)</td>
</tr>
</table>
<!-- inicio aulas -->
<%

datacomp=dtgrade

sql2="select top 100 percent g.chapa1, g.coddoc, e.curso, e.codccusto, e.sal, " & _
"aulas=sum(case when juntar=1 then 0 else case when extra=1 then 0 else case when demons=1 then 0 else 1 end end end), " & _
"e.adnot, sum(g.adnot) noturno, e.aextra, aextras=sum(case when extra=1 and juntar=0 then 1 else 0 end) " & _
"from g2ch g, g2cursoeve e where g.coddoc=e.coddoc and g.deletada=0 and g.inicio<='" & dtaccess(datacomp) & "' and g.demons>=0 " & _
"and '" & dtaccess(datacomp) & "' between g.inicio and g.termino " & _
"group by g.chapa1, g.coddoc, e.curso, e.codccusto, e.sal, e.adnot, e.aextra having g.chapa1='" & rs("chapa") & "' order by e.sal "
rs1.Open sql2, ,adOpenStatic, adLockReadOnly
'response.write "<br>" & rs'.recordcount
if rs1.recordcount>0 then
'rs2.movefirst
do while not rs1.eof
%>
<hr>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=650>
<tr>
	<td class=titulop colspan=7>Curso: <%=rs1("coddoc") & "-" & rs1("curso")%></td>
	<td class=titulop colspan=4>C.Custo: <%=rs1("codccusto")%></td>
</tr>
<%
sql3="SELECT g.perlet, g.chapa1, g.coddoc, c.curso, g.turno, t.descturno, g.codtur, g.diasem, g.DESCRICAO, g.codmat, m.materia, g.juntar, g.dividir, g.extra, g.inicio, g.termino, " & _
"aulas=case when juntar=1 then 0 else 1 end, g.adnot AS noturno " & _
"FROM g2ch AS g, g2cursoeve c, corporerm.dbo.umaterias m, eturnos t " & _
"where g.coddoc=c.coddoc and m.codmat collate database_default=g.codmat and t.codturno=g.turno " & _
"and g.chapa1='" & rs("chapa") & "' AND '" & dtaccess(datacomp) & "' Between g.inicio And g.termino AND c.coddoc='" & rs1("coddoc") & "' AND g.deletada=0 AND g.ativo In (1,0) " & _
"ORDER BY g.chapa1, g.coddoc, g.diasem, g.turno, g.DESCRICAO "
rs2.Open sql3, ,adOpenStatic, adLockReadOnly
%>
<tr>
	<td class=titulor align="center" width=10 >#</td>
	<td class=titulor align="center" width=40 >Turno</td>
	<td class=titulor align="center" width=20 >Dia</td>
	<td class=titulor align="center" width=50 >Turma</td>
	<td class=titulor align="center" width=70 >Horario</td>
	<td class=titulor align="center" width=<%=650-340%> >Materia</td>
	<td class=titulor align="center" width=30 >Junta</td>
	<td class=titulor align="center" width=30 >Divide</td>
	<td class=titulor align="center" width=30 >Extra</td>
	<td class=titulor align="center" width=30 >Ad.Not.</td>
	<td class=titulor align="center" width=30 >Início</td>
</tr>	
<%
rs2.movefirst
do while not rs2.eof
if rs2("noturno")=0 then noturno="&nbsp;" else noturno=rs2("noturno")
if cdate(rs2("inicio"))<>dateserial(2016,8,1) then inicio=rs2("inicio") else inicio="&nbsp;"
%>
<tr>
	<td class="campor"><%=rs2.absoluteposition%></td>
	<td class="campor" nowrap><%=left(rs2("descturno"),3)%></td>
	<td class="campor" align="center"><%=weekdayname(rs2("diasem"),1)%></td>
	<td class="campor" nowrap align="center"><%=rs2("codtur")%></td>
	<td class="campor" align="center"><%=rs2("descricao")%></td>
	<td class="campor"><%=rs2("materia")%></td>
	<td class="campor" align="center"><%if rs2("juntar")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" align="center"><%if rs2("dividir")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" align="center"><%if rs2("extra")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" align="center"><%=noturno%></td>
	<td class="campor" align="center"><%=inicio%></td>
</tr>	
<%
rs2.movenext
loop
rs2.close
%>
</table>
<%
rs1.movenext
loop
end if 'rs1.recordcount
rs1.close
%>
<hr>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=650>
<tr>
	<td class=grupo align="center">Evento</td>
	<td class=grupo align="center">Curso</td>
	<td class=grupo align="center">Aulas</td>
	<td class=grupo align="center">Jornada</td>
	<td class=grupo align="center">Hora</td>
	<td class=grupo align="center">Salário</td>
	<td class=grupo align="center">Ev.Ad.Not.</td>
	<td class=grupo align="center">Hs Ad.Not.</td>
	<td class=grupo align="center">Ev.A.Extra</td>
	<td class=grupo align="center">Hs A.Extra</td>
</tr>	
<%
sql5="select t.sal, last(curso) as curso1, sum(t.aulas) as aulas, t.adnot, sum(t.noturno) as noturno, t.aextra, sum(t.aextras) as aextras " & _
"from (" & sql2 & ") as t " & _
"group by t.sal, t.adnot, t.aextra order by t.sal "
sql5=sql2
rs1.Open sql5, ,adOpenStatic, adLockReadOnly
if rs1.recordcount>0 then
rs1.movefirst
do while not rs1.eof

data=dtsalario
dataref=rs("dataadmissao")
dataref=rs("inicioprofessor")
if dataref<dateserial(2005,2,1) then
	reformulacao="A"
elseif dataref<dateserial(2007,2,1) then
	reformulacao="B"
elseif dataref<dateserial(2009,8,1) then
	reformulacao="C"
elseif dataref<dateserial(2010,2,1) then
	reformulacao="D"
elseif dataref<dateserial(2016,1,1) then
	reformulacao="E"
elseif dataref<dateserial(2016,7,1) then
	reformulacao="F"
else
	reformulacao="G"
end if

sql4="SELECT c.evento, c.tabela, f.dt_faixa, t.titulacao, t.nivel, t.titulo, t.faixasalarial, f.valoraula " & _
"FROM (csd_cursos c INNER JOIN csd_titulos t ON c.tabela=t.tabela) INNER JOIN csd_faixas f ON t.faixasalarial=f.faixasalarial " & _
"WHERE c.evento='" & rs1("sal") & "' " & _
"AND '" & dtaccess(data) & "' Between [ivigencia] And [fvigencia] " & _
"AND f.dt_faixa='" & dtaccess(data) & "' " & _
"AND t.titulacao='" & titulacao & "' AND t.nivel='" & nivel & "' and reformulacao='" & reformulacao & "' "
rs2.Open sql4, ,adOpenStatic, adLockReadOnly
'response.write sql4
if rs2.recordcount>0 then
	titulo=rs2("titulo")
	faixa=rs2("faixasalarial")
	valor=cdbl(rs2("valoraula"))
else
	titulo="":faixa="":valor=0
end if 'recordcount rs4
rs2.close
if rs1("noturno")=0 then evadnot="&nbsp;" else evadnot=rs1("adnot")
if rs1("aextras")=0 then evextra="&nbsp;" else evextra=rs1("aextra")

sqlt="select hativ, dsrem, resol from g2cursoeve where sal='" & rs1("sal") & "'"
rs2.Open sqlt, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	aha=rs2("hativ"):dsr=rs2("dsrem"):res=rs2("resol")
else
	aha="":dsr="":res=""
end if
rs2.close

%>
<tr>
	<td class="campop" align="center"><%=rs1("sal")%></td>
	<td class=campo ><%=rs1("curso")%></td>
	<td class="campop" align="center"><%=rs1("aulas")%></td>
	<td class="campop" align="center"><%=cdbl(rs1("aulas"))*4.5%></td>
	<td class="campop" align="right"><%=formatnumber(valor,2)%>&nbsp;</td>
	<td class="campop" align="right"><%=formatnumber(rs1("aulas")*4.5*valor,2)%>&nbsp;</td>
	<td class="campop" align="center"><%=evadnot%></td>
	<td class="campop" align="center"><%=rs1("noturno")*4.5%></td>
	<td class=campo  align="center"><%=evextra%> <%=aha%>/<%=dsr%>/<%=res%></td>
	<td class="campop" align="center"><%=rs1("aextras")*4.5%></td>
</tr>	
<%
ttaulas=ttaulas + rs1("aulas")
ttjornada=ttjornada + (rs1("aulas")*4.5)
ttsalario=ttsalario + (rs1("aulas")*4.5*valor)
rs1.movenext
loop
end if 'rs2.recordcount

sql4="SELECT sc.CHAPA, sc.CODEVENTO, e.DESCRICAO, sc.NROSALARIO, sc.JORNADA, sc.VALOR, [JORNADA]/60 AS HSMES, [JORNADA]/60/4.5 AS HSSEM, [VALOR]/([JORNADA]/60) AS HORA, sc.INICIOVIGENCIA, sc.FIMVIGENCIA " & _
"FROM corporerm.dbo.PFSALCMP sc INNER JOIN corporerm.dbo.PEVENTO e ON sc.CODEVENTO = e.CODIGO " & _
"WHERE sc.CHAPA='" & rs("chapa") & "' " & _
"AND sc.CODEVENTO IN ('255','256','257','258','128','138') " & _
"ORDER BY sc.CHAPA, sc.NROSALARIO "
'response.write sql4
rs2.Open sql4, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
tsalario=0:thoras=0
rs2.movefirst
do while not rs2.eof
%>
<tr>
	<td class="campop" align="center"><%=rs2("codevento")%></td>
	<td class=campo><%=rs2("descricao")%></td>
	<td class="campop" align="center"><%=rs2("hsmes")/4%></td>
	<td class="campop" align="center"><%=rs2("hsmes")%>&nbsp;</td>
	<td class="campop" align="right"><%=formatnumber(rs2("hora"),2)%>&nbsp;</td>
	<td class="campop" align="right"><%=formatnumber(rs2("valor"),2)%>&nbsp;</td>
	<td class="campop"></td>
	<td class="campop"></td>
	<td class="campop"></td>
	<td class="campop"></td>
</tr>
<%
ttaulas=ttaulas + cdbl((rs2("hsmes")/4))
ttjornada=ttjornada + cdbl(rs2("hsmes"))
ttsalario=ttsalario + cdbl(rs2("valor"))
rs2.movenext
loop
end if 'rs4.recordcount
rs2.close
if ttsalario>0 then valormedio=ttsalario/ttjornada else valormedio=0
%>
<tr>
	<td class=titulo colspan=2></td>
	<td class=titulo align="center"><%=ttaulas%></td>
	<td class=titulo align="center"><%=ttjornada%></td>
	<td class=titulo align="right"><%=formatnumber(valormedio,2)%>&nbsp;</td>
	<td class=titulo align="right"><%=formatnumber(ttsalario,2)%>&nbsp;</td>
	<td class=titulo align="center" colspan=4></td>
</tr>	

</table>
<%
rs1.close

sqln="select CHAPA, n.NOMEACAO, i.CH, i.CODEVE from n_indicacoes i " & _
"inner join n_nomeacoes n on n.id_nomeacao=i.id_nomeacao " & _
"where CHAPA='" & rs("chapa") & "' and GETDATE() between MAND_INI and MAND_FIM"
rs1.Open sqln, ,adOpenStatic, adLockReadOnly
do while not rs1.eof
	response.write "<span style=font-size:14px>" & rs1("nomeacao") & " - " & rs1("codeve") & " - " & rs1("ch") & "</span>"
	response.write "<br>"
rs1.movenext
loop
rs1.close
%>
<%
response.write now()
if rs.absoluteposition<>rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>"
	session("ultimacomposicao")=rs("chapa")	
rs.movenext
loop
rs.close

if teste=1 then
'response.write "Professores lançados no Grades que não tem cadastro ativo no Labore "
sql4="SELECT g.chapa1, f.CHAPA " & _
"FROM g2ch g LEFT JOIN corporerm.dbo.pfunc f ON g.chapa1 = f.CHAPA collate database_default " & _
"WHERE g.perlet Like '2011%2' " & _
"GROUP BY g.chapa1, f.CHAPA " & _
"HAVING f.CHAPA Is Null; "
rs1.Open sql4, ,adOpenStatic, adLockReadOnly
if rs1.recordcount>0 then
rs1.movefirst
do while not rs1.eof
	response.write "<br>" & rs1("chapa1")
rs1.movenext
loop
end if
rs1.close
end if 'teste=1

end if 'finaliza=1


set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>