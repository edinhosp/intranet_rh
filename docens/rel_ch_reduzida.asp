<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a10")="N" or session("a10")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Carga Horária (formato reduzido)</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B1")="" then
%>
<!-- modelo do relatorio inicio -->
<table border=1 bordercolor=#000000 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=670>
<th>Modelo do relatório</th>
<tr><td valign=top>

<p class=realce style="margin-top:0; margin-bottom:0">Carga Horária Resumida em 15/03/04</p>
<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="660">
<tr>
	<td class="titulor">Docente</td>
	<td class="titulor">Dia</td>
	<td class="titulor">Turma</td>
	<td class="titulor">Curso</td>
	<td class="titulor">Disciplina/Atribuição</td>
	<td class="titulor" align="center" colspan=4>Carga Horária</td>
</tr>

<tr>
	<td class="campoa"r colspan="4"><b>ALEXY DUBOIS (00060)</td>
	<td class="campoa"r colspan="1">Adm.: 01/10/76 Titulação: <b>Doutor</b> / CH.T. 28</td>
	<td class="campoa"r align="center">Aulas</td>
	<td class="campoa"r align="center">Adm.</td>
	<td class="campoa"r align="center">Acad.</td>
	<td class="campoa"r align="center">Total</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center">ter</td>
	<td class="campor" align="center">3B</td>
	<td class="campor">ADM. COM  HAB. EM ADMINISTRAÇÃO HOTELEIRA</td>
	<td class="campor">CUSTOS E FORMAÇÃO DE PREÇOS</td>
	<td class="campor" align="center">4</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">4</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center">qui</td>
	<td class="campor" align="center">3C</td>
	<td class="campor">ADM. COM  HAB. EM MARKETING</td>
	<td class="campor">CUSTOS E FORMAÇÃO DE PREÇOS</td>
	<td class="campor" align="center">4</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">4</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center">sex</td>
	<td class="campor" align="center">3C</td>
	<td class="campor">ADMINISTRAÇÃO DE EMPRESAS</td>
	<td class="campor">ADMINISTRAÇÃO DE CUSTOS</td>
	<td class="campor" align="center">2</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">2</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center">seg</td>
	<td class="campor" align="center">3A</td>
	<td class="campor">ADMINISTRAÇÃO DE EMPRESAS</td>
	<td class="campor">ADMINISTRAÇÃO DE CUSTOS</td>
	<td class="campor" align="center">2</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">2</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center">seg</td>
	<td class="campor" align="center">3B</td>
	<td class="campor">ADMINISTRAÇÃO DE EMPRESAS</td>
	<td class="campor">ADMINISTRAÇÃO DE CUSTOS</td>
	<td class="campor" align="center">2</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">2</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center">sex</td>
	<td class="campor" align="center">1A</td>
	<td class="campor">CIÊNCIAS CONTÁBEIS</td>
	<td class="campor">ECONOMIA I</td>
	<td class="campor" align="center">2</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">2</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center">sex</td>
	<td class="campor" align="center">2A</td>
	<td class="campor">CIÊNCIAS CONTÁBEIS</td>
	<td class="campor">ECONOMIA II</td>
	<td class="campor" align="center">2</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">2</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor">CIÊNCIAS CONTÁBEIS</td>
	<td class="campor">Coordenador de Curso</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">10</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">10</td>
</tr>

<tr>
	<td class="titulor" colspan="5" align="left">&nbsp;Total 00060</td>
	<td height=14 class="fundor" align="center" nowrap>18</td>
	<td height=14 class="fundor" align="center" nowrap>10</td>
	<td height=14 class="fundor" align="center" nowrap>0</td>
	<td height=14 class="fundor" align="center">28</td>
</tr>
<tr>
	<td class="campoar" colspan="4"><b>MARCOS RIBEIRO PEREIRA BARRETTO (01106)</td>
	<td class="campoar" colspan="1">Adm.: 02/06/03 Titulação: <b>Doutor</b> / CH.T. 40</td>
	<td class="campoar" align="center">Aulas</td>
	<td class="campoar" align="center">Adm.</td>
	<td class="campoar" align="center">Acad.</td>
	<td class="campoar" align="center">Total</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center">seg</td>
	<td class="campor" align="center">4B</td>
	<td class="campor">SISTEMAS DE INFORMAÇÃO</td>
	<td class="campor">TECNOLOGIAS AVANÇADAS EM INTERNET</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">0</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor">SISTEMAS DE INFORMAÇÃO</td>
	<td class="campor">Coordenador da Pós-Graduação</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">0</td>
</tr>

<tr>
	<td class="campor">&nbsp;</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor"></td>
	<td class="campor">SALARIO - RT 40</td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center"></td>
	<td class="campor" align="center">40</td>
	<td class="campor" align="center">40</td>
</tr>

<tr>
	<td class="titulor" colspan="5" align="left">&nbsp;Total 01106</td>
	<td height=14 class="fundor" align="center" nowrap>0</td>
	<td height=14 class="fundor" align="center" nowrap>0</td>
	<td height=14 class="fundor" align="center" nowrap>40</td>
	<td height=14 class="fundor" align="center">40</td>
</tr>
</table>

</td></tr></table>
<!-- modelo do relatorio final -->

<!-- selecoes -->
<form method="POST" action="rel_ch_reduzida.asp" name="form">
<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=500>
<tr><td valign=top colspan=2>
<p style="margin-bottom: 0" class=realce><b>Seleções para o relatório &quot;Carga Horária (formato reduzido)&quot;</b></p>
</td></tr>
	<tr>
		<td class="campor" nowrap width=150>Data base para o relatório:</td>
		<td>&nbsp;<input type=text name=database value=<%=now%> size=10 class=a></td>
	</tr>
	<tr>
		<td class="titulor" nowrap>Tipo da Seleção</td>
		<td class="titulor">Conteúdo da Seleção</td>
	</tr>
	<tr>
		<td class="campor" nowrap><select size="1" name="selecao" onChange="javascript:submit()">
			<option value="1" <%if request.form("selecao")="1" then response.write "selected"%> >Todos</option>
			<option value="2" <%if request.form("selecao")="2" then response.write "selected"%> >Curso</option>
			<option value="3" <%if request.form("selecao")="3" then response.write "selected"%> >Disciplina</option>
			<option value="4" <%if request.form("selecao")="4" then response.write "selected"%> >Professor</option>
			<option value="5" <%if request.form("selecao")="5" then response.write "selected"%> >Setor</option>
			<option value="6" <%if request.form("selecao")="6" then response.write "selected"%> >Diretor</option>
			<option value="7" <%if request.form("selecao")="7" then response.write "selected"%> >Titulação</option>
			<option value="8" <%if request.form("selecao")="8" then response.write "selected"%> >Carga Horária</option>
			<option value="9" <%if request.form("selecao")="9" then response.write "selected"%> >Especial</option>
		</select>
		</td>
		<td class="campor">
<%
combo=0
select case request.form("selecao")
	case "2" 'curso
		combo=1:sqltemp="SELECT codcur as codigo, curso as descricao FROM grades_2 GROUP BY codcur, curso ORDER BY curso"
		sqltemp="SELECT gc.coddoc as codigo, gc.CURSO as descricao " & _
		"FROM g2ch AS gr INNER JOIN g2cursoeve AS gc ON gr.coddoc=gc.coddoc " & _
		"GROUP BY gc.coddoc, gc.CURSO ORDER BY gc.CURSO; "
	case "3" 'disciplina
		combo=1:sqltemp="SELECT materia as codigo, materia as descricao FROM grades_2 GROUP BY materia ORDER BY materia"
	case "4" 'professor
		combo=1:sqltemp="SELECT g.chapa1 as codigo, f.nome as descricao FROM g2ch g, (select chapa, nome from dc_professor union all select chapa collate database_default, nome collate database_default from grades_novos) as f where g.chapa1 collate database_default=f.chapa GROUP BY g.chapa1, f.nome ORDER BY f.nome "
	case "5" 'setor
		combo=1:sqltemp="select codsecao as codigo, secao as descricao from qry_funcionarios f, grades_chapa g where f.chapa=g.chapa collate database_default group by codsecao, secao "
	case "6" 'diretor
		combo=1:sqltemp="select diretor as codigo, diretor as descricao from grades_per where diretor<>'' group by diretor, diretor "
	case "7" 'titulação
		combo=1:sqltemp="select titulacaopaga as codigo, instrucaomec as descricao from qry_funcionarios d, grades_chapa g where d.chapa=g.chapa collate database_default group by titulacaopaga, instrucaomec "
end select
if combo=1 then
%>
<select size="1" name="cselecao">
<%
rs.Open sqltemp, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
<option value="<%=rs("codigo")%>"><%=rs("descricao")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<%
end if 'selecao combo 1
if request.form("selecao")="8" then
%>
entre <input type="text" name="T1" size="3" value="0">
 e <input type="text" name="T2" size="3" value="99"> horas
<%
end if 'selecao 8
%>
		</td>
	</tr>

<tr><td valign=top colspan=2>
<p><input type="submit" class=button value="Visualizar Relatório" name="B1"></p>
</td></tr>

<tr><td valign=top class=campoe colspan=2>
<p style="margin-top: 0; margin-bottom: 0"><font color="#FF0000">Configure a página do seu navegador (Internet
Explorer, Netscape, Mozilla, etc) no sentido PAISAGEM.</font></p>
</td></tr></table>

</form>
<%
end if  'if do request.form

'**********************************************************

if request.form("B1")<>"" then

filtro="":filtro2="":selecao=""
database=cdate(request.form("database"))
sql1="delete from ttcargahoraria_ch where sessao='" & session("usuariomaster") & "' "
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
sql2="SELECT 1 AS tipoch, g.coddoc, gc.CURSO, f.CODSECAO, f.NOME, g.chapa1 chapa, m.materia collate database_default materia, f.DATAADMISSAO, f.CODSITUACAO, f.TITULACAOPAGA, f.INSTRUCAOMEC, g.turno, g.codtur, g.serie, g.turma, g.diasem, " & _
"min(g.horini) horini, max(g.horfim) horfim, sum(g.ta) AS aulas, g.codmat, g.perlet, diretor_='', g.inicio, g.termino, '" & dtaccess(database) & "' AS [database], '' AS portaria, '' AS obs, " & _
"g.juntar, g.jturma, g.dividir, g.extra, g.demons " & _
"FROM g2ch g, dc_professor AS f, g2cursoeve gc, corporerm.dbo.umaterias m " & _
"WHERE g.chapa1=f.CHAPA collate database_default AND g.coddoc=gc.coddoc and m.codmat collate database_default=g.codmat and '" & dtaccess(database) & "' Between [inicio] And [termino] " & _
"group by g.coddoc, gc.CURSO, f.CODSECAO, f.NOME, g.chapa1, m.materia, f.DATAADMISSAO, f.CODSITUACAO, f.TITULACAOPAGA, f.INSTRUCAOMEC, g.turno, g.codtur, g.serie, g.turma, g.diasem, g.codmat, g.perlet, g.inicio, g.termino, g.juntar, g.jturma, g.dividir, g.extra, g.demons "
'"ORDER BY f.CODSECAO, f.NOME, g.chapa1, g.curso, g.materia; "
sql3="union all "
sql4="SELECT 2 AS tipoch, ni.coddoc, gc.curso, f.CODSECAO, f.NOME, ni.CHAPA, nn.NOMEACAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem, '' as a1,'' as a2, ch=case when ni.codeve is null or ni.codeve='' then 0 else ni.ch end, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), '' as diretor, ni.MAND_INI, ni.MAND_FIM, '" & dtaccess(database) & "' , ni.PORTARIA, ni.CARGO, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons " & _
"FROM n_indicacoes AS ni INNER JOIN dc_professor AS f ON ni.CHAPA=f.CHAPA collate database_default INNER JOIN n_nomeacoes AS nn ON ni.id_nomeacao=nn.id_nomeacao LEFT JOIN g2cursoeve gc ON ni.coddoc=gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [mand_ini] And [mand_fim] " 
'"ORDER BY f.CODSECAO, f.NOME, ni.CHAPA, ni.curso, nn.NOMEACAO; "
sql6="SELECT 3 AS tipoch, g.coddoc, gc.curso, f.CODSECAO, f.NOME, f.CHAPA, g.DESCRICAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem,'' as a1,'' as a2, g.CH, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), null as diretor, g.inicio, g.FIM, '" & dtaccess(database) & "', '' as portaria, '' as obs, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons " & _
"FROM grades_rt AS g INNER JOIN dc_professor AS f ON g.CHAPA=f.CHAPA collate database_default LEFT JOIN g2cursoeve gc ON g.coddoc=gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [inicio] And [fim] " 
'"ORDER BY f.CODSECAO, f.NOME, f.CHAPA, g.curso, g.DESCRICAO; "
'response.write "<br>" & sql2
'response.write "<br>" & sql4
'response.write "<br>" & sql6
sql10=sql2 & sql3 & sql4 & sql3 & sql6
'response.write "<br>" & sql10

select case request.form("selecao")
	case "1" 'todos
		filtrow=""
		filtroh=""
		selecao="Seleção: todos registros"
	case "2" 'curso
		filtrow=""
		filtroh="HAVING coddoc='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes com aulas/atividades no curso: " & request.form("cselecao")
	case "3" 'disciplina
		filtrow=""
		filtroh="HAVING materia='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes com a disciplina: " & request.form("cselecao")
	case "4" 'professor
		filtrow=""
		filtroh="HAVING ss.chapa='" & request.form("cselecao") & "' "
		selecao="Seleção: apenas o docente com a chapa: " & request.form("cselecao")
	case "5" 'setor
		filtrow=""
		filtroh="HAVING codsecao='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes alocados na seção: " & request.form("cselecao")
	case "6" 'diretor
		filtrow=""
		filtroh="HAVING diretor_='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes cujo Diretor do curso é " & request.form("cselecao")
	case "7" 'titulação
		filtrow=""
		filtroh="HAVING titulacaopaga='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes com a titulação: " & request.form("cselecao")
	case "8" 'carga horaria
		valor1=request.form("T1")
		valor2=request.form("T2")
		filtrow=""
		filtroh="HAVING cargahoraria Between " & valor1 & " And " & valor2 & " "
		selecao="Seleção: docentes com carga horária total entre " & valor1 & " And " & valor2 & " horas "
	case "9" 'especial
		filtrow=""
		filtroh="HAVING ss.chapa In (select chapa from zselecao where sessao='" & session("usuariomaster") & "') "
		selecao="Seleção: docentes específicos."
end select

sqla="SELECT CODSECAO, s.descricao as SECAO, NOME, ss.CHAPA, ss.tipoch, coddoc, curso, DATAADMISSAO, CODSITUACAO, titulacaopaga as GRAUINSTRUCAO, instrucaomec as INSTRUCAO, turno, " & _
"codtur, serie, turma, diasem, horini, horfim, codmat, materia, Sum(aulas) AS taulas, perlet, " & _
"diretor_, inicio, termino, ss.[database], juntar, jturma, dividir, extra, demons, PORTARIA, OBS, ch.cargahoraria " & _
"FROM (" & sql10 & ") as ss, ttcargahoraria_ch ch, corporerm.dbo.psecao s "
sqlb="where s.codigo=ss.codsecao and ss.chapa=ch.chapa and ch.sessao='" & session("usuariomaster") & "' and ch.tipoch=4 " & Filtrow
sqlc="GROUP BY CODSECAO, descricao, NOME, ss.CHAPA, ss.tipoch, coddoc, curso, DATAADMISSAO, CODSITUACAO, titulacaopaga, INSTRUCAOmec, turno, " & _
"codtur, serie, turma, diasem, horini, horfim, codmat, materia, perlet, diretor_, inicio, termino, ss.[database], juntar, jturma, dividir, " & _
"extra, demons, PORTARIA, OBS, ch.cargahoraria "
sqld=Filtroh
sqle="ORDER BY nome, CODSECAO, descricao, ss.tipoch, curso, turno, serie, turma, diasem, horini "

sqlz=sqla & sqlb & sqlc & sqld & sqle
'response.write sqlz
rs.Open sqlz, ,adOpenStatic, adLockReadOnly

tfaulas=0:tfadm=0:tfacad=0
tgaulas=0:tgadm=0:tgacad=0
inicio=1:novachapa=1
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
%>
<p class=realce style="margin-top:0; margin-bottom:0">Carga Horária Resumida em <%=database%></p>
<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="660">
<tr>
	<td class="titulor">Dia</td>
	<td class="titulor">Turma</td>
	<td class="titulor">Curso</td>
	<td class="titulor">Disciplina/Atribuição</td>
	<td class="titulor" align="center" colspan=4>Carga Horária</td>
</tr>
<%
linhas=2
rs.movefirst
do while not rs.eof 
chapach=rs("chapa")
session("chapa")=chapach

if linhas>70 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p style='margin-top:0; margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<p class=realce style='margin-top:0; margin-bottom:0'>Carga Horária Resumida em " & database & "</p>"
	linhas=1
	response.write "<table border='1' bordercolor=#000000 cellpadding='1' cellspacing='0' style='border-collapse: collapse' width='660'>"
	response.write "<tr>"
	response.write "<td class=""titulor"">Dia</td>"
	response.write "<td class=""titulor"">Turma</td>"
	response.write "<td class=""titulor"">Curso</td>"
	response.write "<td class=""titulor"">Disciplina/Atribuição</td>"
	response.write "<td class=""titulor"" align=""center"" colspan=4>Carga Horária</td>"
	response.write "</tr>"
	linhas=linhas+1
end if

if inicio=1 then
%>
<tr>
	<td class="campoa"r colspan="3"><b><%=rs("nome") & " (" & rs("chapa") & ")" %></td>
	<td class="campoa"r colspan="1">Adm.: <%=rs("dataadmissao")%> Titulação: <b><%=rs("instrucao")%></b> / CH.T. <%=rs("cargahoraria")%></td>
	<td class="campoa"r align="center">Aulas</td>
	<td class="campoa"r align="center">Adm.</td>
	<td class="campoa"r align="center">Acad.</td>
	<td class="campoa"r align="center">Total</td>
</tr>
<%
	linhas=linhas+1
end if

if inicio=0 then
	if lastchapa<>rs("chapa") then
%>
<tr>
	<td class="titulor" colspan="4" align="left">&nbsp;Total <%=lastchapa %></td>
	<td height=14 class="fundor" align="center" nowrap><%=tfaulas%></td>
	<td height=14 class="fundor" align="center" nowrap><%=tfadm%></td>
	<td height=14 class="fundor" align="center" nowrap><%=tfacad %></td>
	<td height=14 class="fundor" align="center"><%=tfaulas+tfacad+tfadm %></td>
</tr>
<tr>
	<td class="campoa"r colspan="3" nowrap><b><%=rs("nome") & " (" & rs("chapa") & ")" %></td>
	<td class="campoa"r colspan="1" nowrap>Adm.: <%=rs("dataadmissao")%> Titulação: <b><%=rs("instrucao")%></b> / CH.T. <%=rs("cargahoraria")%></td>
	<td class="campoa"r align="center">Aulas</td>
	<td class="campoa"r align="center">Adm.</td>
	<td class="campoa"r align="center">Acad.</td>
	<td class="campoa"r align="center">Total</td>
</tr>
<%
	linhas=linhas+2
	novachapa=1
	tfaulas=0:tfacad=0:tfadm=0
	end if
end if
if rs("tipoch")=1 then var1=rs("taulas") else var1=0
if rs("tipoch")=2 then var2=rs("taulas") else var2=0
if rs("tipoch")=3 then var3=rs("taulas") else var3=0
if rs("diasem")<>0 then diasem=weekdayname(rs("diasem"),1) else diasem=""
'if rs("horaini")="" or isnull(rs("horaini")) then horini="" else horini=formatdatetime(rs("horaini"),4)
'if rs("horafim")="" or isnull(rs("horafim")) then horfim="" else horfim=formatdatetime(rs("horafim"),4)
'if horini<>"" and horfim<>"" then horario=horini & " a " & horfim else horario=""
if rs("curso")="" then complemento=" / " & rs("obs") else complemento=""
%>
<tr>
	<td class="campor" align="center"><%=diasem%></td>
	<td class="campor" align="center" nowrap><%=rs("codtur")%></td>
	<td class="campor" nowrap><%=left(rs("curso"),35)%></td>
	<td class="campor" nowrap><%=left(rs("materia"),40)%></td>
	<td class="campor" align="center"><%if var1<>0 then response.write var1%></td>
	<td class="campor" align="center"><%if var2<>0 then response.write var2%></td>
	<td class="campor" align="center"><%if var3<>0 then response.write var3%></td>
	<td class="campor" align="center"><%=rs("taulas")%></td>
</tr>
<%
linhas=linhas+1
inicio=0
lastsecao=rs("codsecao")
lastchapa=rs("chapa")
tfaulas =tfaulas + var1
tfadm   =tfadm   + var2
tfacad  =tfacad  + var3
tgaulas =tgaulas + var1
tgadm   =tgadm   + var2
tgacad  =tgacad  + var3

rs.movenext
loop
rs.close
set rs=nothing
%>
<tr>
	<td class="titulor" colspan="4" align="left">&nbsp;Total <%=lastchapa %></td>
	<td height=14 class="fundor" align="center" nowrap><%=tfaulas%></td>
	<td height=14 class="fundor" align="center" nowrap><%=tfadm%></td>
	<td height=14 class="fundor" align="center" nowrap><%=tfacad%></td>
	<td height=14 class="fundor" align="center"><%=tfaulas+tfacad+tfadm%></td>
</tr>
<tr><td class="campor" colspan="14"><hr></td></tr>
<tr>
	<td class="titulor" colspan="4">&nbsp;Total Geral</td>
	<td height=14 class="campolr" align="center" nowrap><%=tgaulas%></font></td>
	<td height=14 class="campolr" align="center" nowrap><%=tgadm  %></font></td>
	<td height=14 class="campolr" align="center" nowrap><%=tgacad %></font></td>
	<td height=14 class="campolr" align="center"><%=tgaulas+tgacad+tgadm%></font></td>
</tr>
</table>
<p><i><font size="1" color="#0000FF"><b><%=selecao %></b></font></i></p>
<%	pagina=pagina+1
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
%>
<%
else 'sem registros
%>
<p>
<b><font color="#FF0000">
Esta seleção não mostra nenhum registro.</font></b></p>
<%
end if 'recordcount

end if 'if do request.form
%>
<%
conexao.close
set conexao=nothing
%>
</body>
</html>