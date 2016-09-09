<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a15")="N" or session("a15")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Distribuição de docentos pela grade</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
	dim conexao, conexao2, chapach, rs, rs2, tgl(5,6), tl(5), tg(6), descricao(5)
	set conexao=server.createobject ("ADODB.Connection")
	conexao.Open application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	Set rs.ActiveConnection = conexao

filtro="":filtro2="":selecao=""
database="15/06/2012"
database=formatdatetime(now(),2)

sql1="delete from ttcargahoraria_ch where sessao='" & session("usuariomaster") & "' ":conexao.execute sql1
'************ total
sql2="INSERT INTO ttcargahoraria_ch ( sessao, tipoch, CHAPA, cargahoraria, [database] ) " & _
"SELECT '" & session("usuariomaster") & "', 1 , chapa1, Sum(ta), '" & dtaccess(database) & "' " & _
"FROM g2ch WHERE '" & dtaccess(database) & "' Between [inicio] And [termino] GROUP BY chapa1 "
'response.write "<br>" & sql2
sql3="INSERT INTO ttcargahoraria_ch ( sessao, tipoch, CHAPA, cargahoraria, [database] ) " & _
"SELECT '" & session("usuariomaster") & "', 2 , CHAPA, sum(case when codeve is null or codeve='' then 0 else ch end), '" & dtaccess(database) & "' " & _
"FROM n_indicacoes WHERE '" & dtaccess(database) & "' Between [mand_ini] And [mand_fim] GROUP BY CHAPA "
'response.write "<br>" & sql3
sql4="INSERT INTO ttcargahoraria_ch ( sessao, tipoch, CHAPA, cargahoraria, [database] ) " & _
"SELECT '" & session("usuariomaster") & "', 3 , CHAPA, Sum(CH), '" & dtaccess(database) & "' " & _
"FROM grades_rt WHERE '" & dtaccess(database) & "' Between [inicio] And [fim] GROUP BY CHAPA "
'response.write "<br>" & sql4
sql5="INSERT INTO ttcargahoraria_ch ( sessao, tipoch, CHAPA, cargahoraria, [database] ) " & _
"SELECT sessao, 4, CHAPA, Sum(cargahoraria) AS SomaDecargahoraria, '" & dtaccess(database) & "' " & _
"FROM ttcargahoraria_ch GROUP BY sessao, CHAPA, [database] HAVING sessao='" & session("usuariomaster") & "' and [database]='" & dtaccess(database) & "' "
'response.write "<br>" & sql5
sql12="SELECT sessao, [database] FROM ttcargahoraria_ch GROUP BY sessao, [database] HAVING sessao='" & session("usuariomaster") & "' " & _
"and [database]='" & dtaccess(database) & "' "
'response.write "<br>" & sql12
rs.Open sql12, ,adOpenStatic, adLockReadOnly
if rs.recordcount=0 then
	conexao.execute sql1
	conexao.execute sql2:conexao.execute sql3:conexao.execute sql4:conexao.execute sql5
end if
rs.close
'********s*sassssdaaas**************
sql2="SELECT 1 AS tipoch, g.coddoc, gc.CURSO, f.CODSECAO, f.NOME, g.chapa1 collate database_default as chapa, g.materia, f.DATAADMISSAO, f.CODSITUACAO, f.TITULACAOPAGA, " & _
"f.INSTRUCAOMEC, g.turno, g.codtur, g.serie, g.turma, g.diasem, g.a1, g.a2, g.a3, g.a4, g.a5, g.a6, g.ta AS aulas, g.codmat, gp.perlet, gp.perlet2, " & _
"diretor_=case when diretor is null and perlet3>='2004' and turno in (1,2,5) then 'Maria Celia Soares Hungria de Luca' else (case when diretor is null and perlet3>='2004' and turno in (3) then 'Luiz Carlos de Azevedo Filho' else diretor end) end, " & _
"g.inicio, g.termino, '" & dtaccess(database) & "' AS [database], '' AS portaria, '' AS obs, g.juntar, g.jturma, g.dividir, g.extra, g.demons " & _
"FROM ((g2ch AS g INNER JOIN dc_professor AS f ON g.chapa1 collate database_default=f.CHAPA collate database_default) INNER JOIN " & _
"(select coddoc, perlet, perlet2, enfase, pini, pfim, lanc, diretor from grades_per group by coddoc, perlet, perlet2, enfase, pini, pfim, lanc, diretor) " & _
"AS gp ON (gp.enfase = g.enfase) AND (g.perlet2 = gp.perlet2) AND (g.perlet = gp.perlet) AND (g.coddoc = gp.coddoc)) INNER JOIN g2cursoeve AS gc ON g.coddoc = gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [inicio] And [termino] "
'"ORDER BY f.CODSECAO, f.NOME, g.chapa1, g.curso, g.materia; "
sql3="union all "
sql4="SELECT 2 AS tipoch, ni.coddoc, gc.curso, f.CODSECAO, f.NOME, ni.CHAPA, nn.NOMEACAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem,'' as a1,'' as a2,'' as a3,'' as a4,'' as a5,'' as a6, ch=case when ni.codeve is null or ni.codeve='' then 0 else ni.ch end, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), cast(Year(getdate()) as char(4)), '' as diretor, ni.MAND_INI, ni.MAND_FIM, '" & dtaccess(database) & "' , ni.PORTARIA, ni.CARGO, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons " & _
"FROM n_indicacoes AS ni INNER JOIN dc_professor AS f ON ni.CHAPA=f.CHAPA collate database_default INNER JOIN n_nomeacoes AS nn ON ni.id_nomeacao=nn.id_nomeacao LEFT JOIN g2cursoeve gc ON ni.coddoc=gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [mand_ini] And [mand_fim] " 
'"ORDER BY f.CODSECAO, f.NOME, ni.CHAPA, ni.curso, nn.NOMEACAO; "
sql6="SELECT 3 AS tipoch, g.coddoc, gc.curso, f.CODSECAO, f.NOME, f.CHAPA, g.DESCRICAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem,'' as a1,'' as a2,'' as a3,'' as a4,'' as a5,'' as a6, g.CH, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), cast(Year(getdate()) as char(4)), null as diretor, g.inicio, g.FIM, '" & dtaccess(database) & "', '' as portaria, '' as obs, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons " & _
"FROM grades_rt AS g INNER JOIN dc_professor AS f ON g.CHAPA=f.CHAPA collate database_default LEFT JOIN g2cursoeve gc ON g.coddoc=gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [inicio] And [fim] " 
'"ORDER BY f.CODSECAO, f.NOME, f.CHAPA, g.curso, g.DESCRICAO; "
'response.write "<br>" & sql2
'response.write "<br>" & sql4
'response.write "<br>" & sql6
sql10=sql2 & sql3 & sql4 & sql3 & sql6
'response.write "<br>" & sql10

sql1="select particao=case when cargahoraria<12 then '1-Horista (até 12 horas)' else " & _
"case when cargahoraria<=29 then '2-Parcial (entre 12 e 29 horas)' else " & _
"case when cargahoraria<=39 then '3-Parcial (entre 30 e 39 horas)' else '4-Integral (40 horas)' end end end, " & _
"sum(case f.instrucaomec when 'Graduado' then 1 else 0 end) as 'Educação', " & _
"sum(case f.instrucaomec when 'Especialista' then 1 else 0 end) as 'Especialista', " & _
"sum(case f.instrucaomec when 'Mestrando' then 1 else 0 end) as 'Mestrando', " & _
"sum(case f.instrucaomec when 'Doutorando' then 1 else 0 end) as 'Doutorando', " & _
"sum(case f.instrucaomec when 'Mestre' then 1 else 0 end) as 'Mestre', " & _
"sum(case f.instrucaomec when 'Doutor' then 1 else 0 end) as 'Doutor' " & _
"FROM ttcargahoraria_ch t INNER JOIN dc_professor f ON t.CHAPA=f.CHAPA collate database_default " & _
"WHERE t.sessao='" & session("usuariomaster") & "' AND t.tipoch=4 and t.cargahoraria>0 /*and f.codsituacao in ('A','F','Z','E')*/  " & _
"and dataadmissao<='" & dtaccess(database) & "' and (datademissao is null or datademissao>'" & dtaccess(database) & "') " & _
"group by case when cargahoraria<12 then '1-Horista (até 12 horas)' else " & _
"case when cargahoraria<=29 then '2-Parcial (entre 12 e 29 horas)' else " & _
"case when cargahoraria<=39 then '3-Parcial (entre 30 e 39 horas)' else '4-Integral (40 horas)' end end end "

'response.write "<br>" & sql1
rs.Open sql1, ,adOpenStatic, adLockReadOnly
'response.write "<br>" & rs.recordcount

tfaulas=0:tfadm=0:tfacad=0
tgaulas=0:tgadm=0:tgacad=0
inicio=1
'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="titulor">" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
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

if rs.recordcount>0 then 
rs.movefirst
do while not rs.eof
	for a=1 to 6
		if rs.fields(a)<>"" then
			tgl(rs.absoluteposition,a)=cdbl(rs.fields(a))
			tl(rs.absoluteposition)=tl(rs.absoluteposition) + cdbl(rs.fields(a))
			tg(a)=tg(a)+cdbl(rs.fields(a))
			total=total+cdbl(rs.fields(a))
		end if
	next
	descricao(rs.absoluteposition)=rs.fields(0)
rs.movenext
loop
rs.close
%>
<p class=titulo>Distribuição de docentes por titulação
<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulo align="center" rowspan=2 style="border-right:2px solid">Distribuição</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Graduado</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Especialista</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Mestrando</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Doutorando</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Mestre</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Doutor</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Totais</td>
</tr>
<tr>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
</tr>
<%
for a=1 to 4 '4
%>
<tr>
	<td class=campo style="border-right:2px solid">&nbsp;<%=descricao(a)%></td>
<%
	for b=1 to 6
%>
	<td class="campoa" align="center"><b><%=tgl(a,b)%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((tgl(a,b)/total)*100,1)%></td>
<%
	next
%>
	<td class="campol" align="center"><b><%=tl(a)%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((tl(a)/total)*100,1)%></td>
</tr>
<%
next  
%>
<tr>
	<td class=titulo style="border-right:2px solid">&nbsp;Totais</td>
<%
	for a=1 to 6
%>
	<td class="campot" align="center"><b><%=tg(a)%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((tg(a)/total)*100,1)%></td>
<%
	next
%>
	<td class=titulop align="center"><%=total%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((total/total)*100,1)%></td>
</tr>
</table>

<!-- RESUMOS -->

<p class=titulo>Distribuição de docentes por titulação (Resumo)
<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulo align="center" rowspan=2 style="border-right:2px solid">Distribuição</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Grad/Esp/Outros</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Mestre/Doutor</td>
	<td class=titulo align="center" colspan=2 style="border-right:2px solid">Totais</td>
</tr>
<tr>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center" style="border-right:2px solid">%</td>
</tr>
<%
for a=1 to 4' 4
%>
<tr>
	<td class=campo style="border-right:2px solid">&nbsp;<%=descricao(a)%></td>
<%
tt1=tgl(a,1)+tgl(a,2)+tgl(a,3)+tgl(a,4)
tt2=tgl(a,5)+tgl(a,6)
tt1=tgl(a,1)+tgl(a,2)+tgl(a,3)
tt2=tgl(a,5)+tgl(a,6)+tgl(a,4)

if a=3 then ttt=tt2
%>
	<td class="campoa" align="center"><b><%=tt1%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((tt1/total)*100,1)%></td>
	<td class="campoa" align="center"><b><%=tt2%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((tt2/total)*100,1)%></td>
<%
%>
	<td class="campol" align="center"><b><%=tl(a)%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((tl(a)/total)*100,1)%></td>
</tr>
<%
next  
%>
<tr>
	<td class=titulo style="border-right:2px solid">&nbsp;Totais</td>
<%
ttg1=tg(1)+tg(2)+tg(3)+tg(4)
ttg2=tg(5)+tg(6)
ttg1=tg(1)+tg(2)+tg(3)
ttg2=tg(5)+tg(6)+tg(4)
if tl(4)-int((total*variavel+0.005)/100)>0 then texto1="Superavit" else texto1="Defícit"
if (tl(3)+tl(4))-int((total*variavel+0.005)/100)>0 then texto2="Superavit" else texto2="Defícit"
%>
	<td class="campot" align="center"><b><%=ttg1%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((ttg1/total)*100,1)%></td>
	<td class="campot" align="center"><b><%=ttg2%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((ttg2/total)*100,1)%></td>
<%
%>
	<td class=titulop align="center"><%=total%></td>
	<td class=campo align="center" style="border-right:2px solid"><%=formatnumber((total/total)*100,1)%></td>
</tr>
</table>

<p class=titulo>Metas da Portaria do MEC para Dezembro/2007
<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=650 >
<tr>
	<td class=titulo align="center" colspan=2>Esperado</td>
	<td class=titulo align="center" colspan=2>Atual</td>
	<td class=titulo align="center" rowspan=2><%=texto1%> (1)</td>
	<td class=titulo align="center" colspan=2>Aum.CH</td>
	<td class=titulo align="center" rowspan=2><%=texto2%> (2)</td>
</tr>
<tr>
	<td class=titulo align="center">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center">%</td>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center">%</td>
	<td class=titulo align="center">Nº</td>
</tr>
<tr>
<%
variavel=20 'anterior 15
%>
	<td class=campo align="center"><%=variavel%>%</td>
	<td class=campo align="center"><%=int((total*variavel+0.005)/100)%></td>
	<td class=campo align="center"><%=formatnumber((tl(4)/total)*100,1)%>%</td>
	<td class=campo align="center"><%=tl(4)%></td>
	<td class=campo align="center"><%=tl(4)-int((total*variavel+0.005)/100)%></td>
	<td class=campo align="center"><%=formatnumber(((tl(3)+tl(4))/total)*100,1)%>%</td>
	<td class=campo align="center"><%=tl(3)+tl(4)%></td>
	<td class=campo align="center"><%=(tl(3)+tl(4))-int((total*variavel+0.005)/100)%></td>
</tr>
</table>
<p><font size=1>(1) Defícit em relação à portaria na situação atual.
<br>(2) Defícit em relação à portaria aumentando a carga horária dos professores na faixa de 30 a 39 horas.
<br>
<br>
<table border="1" bordercolor=#A9A9A9 cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=650 >
<tr>
	<td class=campo valign=top><b>Titulação Paga/Carreira
		<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse" >
		<tr><td class=titulo>Titulação</td><td class=titulo>Quant.</td></tr>	
<%
total1=0:total2=0
sqltp="SELECT f.TITULACAOPAGA, f.INSTRUCAOMEC, Count(t.CHAPA) AS Freq " & _
"FROM ttcargahoraria_ch AS t INNER JOIN dc_professor AS f ON t.CHAPA = f.CHAPA collate database_default " & _
"WHERE t.sessao='" & session("usuariomaster") & "' AND t.tipoch=4 and t.cargahoraria>0 " & _
"and dataadmissao<='" & dtaccess(database) & "' and (datademissao is null or datademissao>'" & dtaccess(database) & "') " & _
"GROUP BY f.TITULACAOPAGA, f.INSTRUCAOMEC;"
rs.Open sqltp, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
		<tr><td class=campo><%=rs("instrucaomec")%></td><td class=campo align="right"><%=rs("freq")%></td><tr>
<%
total1=total1+rs("freq")
rs.movenext
loop
rs.close
%>	
		<tr><td class=titulo>Total</td><td class=titulo align="right"><%=total1%></td></tr>	
		</table>
	</td>
<!-- -->
	<td class=campo valign=top><b>Titulação Real
		<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse" >
		<tr><td class=titulo>Titulação</td><td class=titulo>Quant.</td></tr>	
<%
sqltr="SELECT f.GRAUINSTRUCAO, f.INSTRUCAO, Count(t.CHAPA) AS Freq " & _
"FROM ttcargahoraria_ch AS t INNER JOIN dc_professor AS f ON t.CHAPA = f.CHAPA collate database_default " & _
"WHERE t.sessao='" & session("usuariomaster") & "' AND t.tipoch=4 and t.cargahoraria>0  " & _
"and dataadmissao<='" & dtaccess(database) & "' and (datademissao is null or datademissao>'" & dtaccess(database) & "') " & _
"GROUP BY f.GRAUINSTRUCAO, f.INSTRUCAO;"
rs.Open sqltr, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
		<tr><td class=campo><%=rs("instrucao")%></td><td class=campo align="right"><%=rs("freq")%></td><tr>
<%
total2=total2+rs("freq")
rs.movenext
loop
rs.close
%>	
		<tr><td class=titulo>Total</td><td class=titulo align="right"><%=total2%></td></tr>	
		</table>
	</td>
<!-- -->
	<td class=campo valign=top><b>Tempo de Serviço
		<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse" >
		<tr><td class=titulo>Tempo de Serviço</td><td class=titulo>Quant.</td></tr>	
<%
sqlts="SELECT tempo=" & _
"case when datediff(yy,dataadmissao,getdate())<1 then 'Até 01 ano' else " & _
"case when datediff(yy,dataadmissao,getdate())<=5 then 'De 01 a 05 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=10 then 'De 06 a 10 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=15 then 'De 11 a 15 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=20 then 'De 16 a 20 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=25 then 'De 21 a 25 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=30 then 'De 26 a 30 anos' else " & _
"'Mais de 30 anos' end end end end end end end " & _
", count(f.chapa) as total " & _
"FROM ttcargahoraria_ch AS t INNER JOIN dc_professor AS f ON t.CHAPA = f.CHAPA collate database_default " & _
"WHERE t.sessao='" & session("usuariomaster") & "' AND t.tipoch=4 and t.cargahoraria>0 " & _
"and dataadmissao<='" & dtaccess(database) & "' and (datademissao is null or datademissao>'" & dtaccess(database) & "') " & _
"GROUP BY " & _
"case when datediff(yy,dataadmissao,getdate())<1 then 'Até 01 ano' else " & _
"case when datediff(yy,dataadmissao,getdate())<=5 then 'De 01 a 05 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=10 then 'De 06 a 10 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=15 then 'De 11 a 15 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=20 then 'De 16 a 20 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=25 then 'De 21 a 25 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=30 then 'De 26 a 30 anos' else " & _
"'Mais de 30 anos' end end end end end end end " 
'------------------ para consultoria
sqlts="SELECT tempo=" & _
"case when datediff(yy,dataadmissao,getdate())<1 then 'Até 01 ano' else " & _
"case when datediff(yy,dataadmissao,getdate())<=4 then 'De 01 a 04 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=9 then 'De 05 a 09 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=14 then 'De 10 a 14 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=19 then 'De 15 a 19 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=24 then 'De 20 a 24 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=29 then 'De 25 a 29 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=34 then 'De 30 a 34 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=39 then 'De 35 a 39 anos' else " & _
"'Mais de 40 anos' end end end end end end end end end " & _
", count(f.chapa) as total " & _
"FROM ttcargahoraria_ch AS t INNER JOIN dc_professor AS f ON t.CHAPA = f.CHAPA collate database_default " & _
"WHERE t.sessao='" & session("usuariomaster") & "' AND t.tipoch=4 and t.cargahoraria>0 " & _
"and dataadmissao<='" & dtaccess(database) & "' and (datademissao is null or datademissao>'" & dtaccess(database) & "') " & _
"GROUP BY " & _
"case when datediff(yy,dataadmissao,getdate())<1 then 'Até 01 ano' else " & _
"case when datediff(yy,dataadmissao,getdate())<=4 then 'De 01 a 04 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=9 then 'De 05 a 09 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=14 then 'De 10 a 14 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=19 then 'De 15 a 19 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=24 then 'De 20 a 24 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=29 then 'De 25 a 29 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=34 then 'De 30 a 34 anos' else " & _
"case when datediff(yy,dataadmissao,getdate())<=39 then 'De 35 a 39 anos' else " & _
"'Mais de 40 anos' end end end end end end end end end " 
rs.Open sqlts, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
		<tr><td class=campo><%=rs("tempo")%></td><td class=campo align="right"><%=rs("total")%></td><tr>
<%
total3=total3+rs("total")
rs.movenext
loop
rs.close
%>	
		<tr><td class=titulo>Total</td><td class=titulo align="right"><%=total3%></td></tr>	
		</table>
	</td>
</tr>
</table>

<%
	pagina=pagina+1
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
else 'sem registros
%>
<p>
<b><font color="#FF0000">
Esta seleção não mostra nenhum registro.</font></b></p>
<%
end if 'recordcount

'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>