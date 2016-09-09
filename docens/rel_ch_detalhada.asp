<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a9")="N" or session("a9")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"

%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Carga Horária detalhada</title>
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
<table border=1 bordercolor=#000000 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=950>
<th>Modelo do relatório</th>
<tr><td valign=top>

<p class=realce style="margin-top:0; margin-bottom:0">Carga Horária em 15/03/04</p>
<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="950">

<tr><td height=14 class="grupo" colspan="14">01.3.010 - CURSO DE DIREITO</td></tr>
<tr>
	<td height=14 class="titulor" align="center">Chapa</td>
	<td height=14 class="titulor" align="center">Nome</td>
	<td height=14 class="titulor" align="center">R.T.</td>
	<td height=14 class="titulor" align="center">Titulação</td>
	<td height=14 class="titulor" align="center">Admissão </td>
	<td height=14 class="titulor" align="center">Dia      </td>
	<td height=14 class="titulor" align="center">Horário  </td>
	<td height=14 class="titulor" align="center">Turma    </td>
	<td height=14 class="titulor" align="center">Curso    </td>
	<td height=14 class="titulor" align="center">Disciplina</td>
	<td height=14 class="titulor" align="center">Aulas/CH </td>
	<td height=14 class="titulor" align="center" colspan=3>Observações</td>
</tr>

<tr>
	<td height=28 class="campor" >00590</td>
	<td height=28 class="campor" ><b>ELIANA APARECIDA SANTOS</td>
	<td height=28 class="campor" align="center">8</td>
	<td height=28 class="campor" >Doutorando</td>
	<td height=28 class="campor" align="center">01/07/98</td>
	<td height=28 class="campor" align="center">ter</td>
	<td height=28 class="campor" nowrap align="center">19:30 a 21:10</td>
	<td height=28 class="campor" align="center">4A</td>
	<td height=28 class="campor" >ADMINISTRAÇÃO DE EMPRESAS</td>
	<td height=28 class="campor" >DIREITO TRIBUTÁRIO</td>
	<td height=28 class="campor" align="center">2</td>
	<td height=28 class="campor" colspan=3></td>
</tr>

<tr>
	<td height=28 class="campor" ></td>
	<td height=28 class="campor" ><b></td>
	<td height=28 class="campor" align="center"></td>
	<td height=28 class="campor" ></td>
	<td height=28 class="campor" align="center"></td>
	<td height=28 class="campor" align="center">ter</td>
	<td height=28 class="campor" nowrap align="center">21:20 a 23:00</td>
	<td height=28 class="campor" align="center">4B</td>
	<td height=28 class="campor" >ADMINISTRAÇÃO DE EMPRESAS</td>
	<td height=28 class="campor" >DIREITO TRIBUTÁRIO</td>
	<td height=28 class="campor" align="center">2</td>
	<td height=28 class="campor" colspan=3></td>
</tr>

<tr>
	<td height=14 class="titulor" colspan="10" align="left">Total 00590</td>
	<td height=14 class="campor" align="center">4</td>
	<td height=14 class="campor" align="center" nowrap>Grad: <b>4</td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Adm: <b>0</td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Acad: <b>0</td>
</tr>

<tr>
	<td height=28 class="campor" >00455</td>
	<td height=28 class="campor" ><b>LUIZ FRANCISCO LIPPO</td>
	<td height=28 class="campor" align="center">20</td>
	<td height=28 class="campor" >Mestre</td>
	<td height=28 class="campor" align="center">16/02/96</td>
	<td height=28 class="campor" align="center">seg</td>
	<td height=28 class="campor" nowrap align="center">07:30 a 08:20</td>
	<td height=28 class="campor" align="center">4C</td>
	<td height=28 class="campor" >ADMINISTRAÇÃO DE EMPRESAS</td>
	<td height=28 class="campor" >DIREITO TRIBUTÁRIO</td>
	<td height=28 class="campor" align="center">0</td>
	<td height=28 class="campor" colspan=3></td>
</tr>

<tr>
	<td height=28 class="campor" ></td>
	<td height=28 class="campor" ><b></td>
	<td height=28 class="campor" align="center"></td>
	<td height=28 class="campor" ></td>
	<td height=28 class="campor" align="center"></td>
	<td height=28 class="campor" align="center">seg</td>
	<td height=28 class="campor" nowrap align="center">08:20 a 09:10</td>
	<td height=28 class="campor" align="center">4C</td>
	<td height=28 class="campor" >ADMINISTRAÇÃO DE EMPRESAS</td>
	<td height=28 class="campor" >DIREITO TRIBUTÁRIO</td>
	<td height=28 class="campor" align="center">1</td>
	<td height=28 class="campor" colspan=3></td>
</tr>

<tr>
	<td height=14 class="titulor" colspan="10" align="left">Total 00455</td>
	<td height=14 class="campor" align="center">1</td>
	<td height=14 class="campor" align="center" nowrap>Grad: <b>1</td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Adm: <b>0</td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Acad: <b>0</td>
</tr>

<tr><td height=14 class="grupo" colspan="14">03.3.509 - CURSO DE POS GRADUACAO</td></tr>
<tr>
	<td height=14 class="titulor" align="center">Chapa</td>
	<td height=14 class="titulor" align="center">Nome</td>
	<td height=14 class="titulor" align="center">R.T.</td>
	<td height=14 class="titulor" align="center">Titulação</td>
	<td height=14 class="titulor" align="center">Admissão </td>
	<td height=14 class="titulor" align="center">Dia      </td>
	<td height=14 class="titulor" align="center">Horário  </td>
	<td height=14 class="titulor" align="center">Turma    </td>
	<td height=14 class="titulor" align="center">Curso    </td>
	<td height=14 class="titulor" align="center">Disciplina</td>
	<td height=14 class="titulor" align="center">Aulas/CH </td>
	<td height=14 class="titulor" align="center" colspan=3>Observações</td>
</tr>

<tr>
	<td height=28 class="campor" >00922</td>
	<td height=28 class="campor" ><b>JOSE GERALDO DE LIMA JUNIOR</td>
	<td height=28 class="campor" align="center">40</td>
	<td height=28 class="campor" >Mestre</td>
	<td height=28 class="campor" align="center">01/04/99</td>
	<td height=28 class="campor" align="center">qui</td>
	<td height=28 class="campor" nowrap align="center">07:30 a 11:00</td>
	<td height=28 class="campor" align="center">1D</td>
	<td height=28 class="campor" >ADMINISTRAÇÃO DE EMPRESAS</td>
	<td height=28 class="campor" >TEORIA DA ADMINISTRAÇÃO</td>
	<td height=28 class="campor" align="center">0</td>
	<td height=28 class="campor" colspan=3></td>
</tr>

<tr>
	<td height=14 class="titulor" colspan="10" align="left">Total 00922</td>
	<td height=14 class="campor" align="center">0</td>
	<td height=14 class="campor" align="center" nowrap>Grad: <b>0</td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Adm: <b>0</td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Acad: <b>0</td>
</tr>
</table>

</td></tr></table>
<!-- modelo do relatorio final -->

<!-- selecoes -->
<form method="POST" name="form" action="rel_ch_detalhada.asp">
<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=500>
<tr><td valign=top colspan=2>
	<p style="margin-bottom: 0" class=realce><b>Seleções para o relatório &quot;Carga Horária Detalhada&quot;</b></p>
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
	conexao.execute sql2:
	conexao.execute sql3:
	conexao.execute sql4:
	conexao.execute sql5
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
'response.write "<br><br>" & sqlz
rs.Open sqlz, ,adOpenStatic, adLockReadOnly

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
%>
<p class=realce style="margin-top:0; margin-bottom:0">Carga Horária em <%=database%></p>
<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="950">
<%
linhas=linhas+1
rs.movefirst
do while not rs.eof 
cargahoraria=rs("cargahoraria")
if rs("turno")="1" then turno="Mat"
if rs("turno")="2" then turno="Vesp"
if rs("turno")="3" then turno="Not"
if rs("turno")="5" then turno="Vesp-EF"

if linhas>38 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p style='margin-top:0; margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<p class=realce style='margin-top:0; margin-bottom:0'>Carga Horária em " & database & "</p>"
	linhas=1
	response.write "<table border='1' bordercolor=#000000 cellpadding='1' cellspacing='0' style='border-collapse: collapse' width='950'>"
	response.write "<tr>"
	response.write "<td height=14 class=""titulor"" align=""center"">Chapa</td>"
	response.write "<td height=14 class=""titulor"" align=""center"">Nome</td>"
	response.write "<td height=14 class=""titulor"" align=""center"">R.T.</td>"
	response.write "<td height=14 class=""titulor"" align=""center"">Titulação</td>"
	response.write "<td height=14 class=""titulor"" align=""center"">Admissão </td>"
	response.write "<td height=14 class=""titulor"" align=""center"">Dia      </td>"
	response.write "<td height=14 class=""titulor"" align=""center"">Turno    </td>"
	response.write "<td height=14 class=""titulor"" align=""center"" colspan=2>Horário  </td>"
	response.write "<td height=14 class=""titulor"" align=""center"">Turma    </td>"
	response.write "<td height=14 class=""titulor"" align=""center"">Curso    </td>"
	response.write "<td height=14 class=""titulor"" align=""center"">Disciplina</td>"
	response.write "<td height=14 class=""titulor"" align=""center"">Aulas/CH </td>"
	response.write "<td height=14 class=""titulor"" align=""center"" colspan=3>Observações</td>"
	response.write "</tr>"
	linhas=linhas+1
end if

chapach=rs("chapa")
session("chapa")=chapach
if inicio=0 then
	if lastchapa=rs("chapa") then
	else
%>
<tr>
	<td height=14 class="titulor" colspan="12" align="left">Total <%=lastchapa %></td>
	<td height=14 class="campor" align="center"><%=tfaulas+tfacad+tfadm%></td>
	<td height=14 class="campor" align="center" nowrap>Grad: <b><%=tfaulas%></td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Adm: <b><%=tfadm%></td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Acad: <b><%=tfacad%></td>
</tr>
<%
	linhas=linhas+1
	tfaulas=0
	tfacad=0
	tfadm=0
	end if
end if
if lastsecao=rs("codsecao") then
else
%>
<tr><td height=14 class="grupo" colspan="16"><%=rs("codsecao") & " - " & rs("secao") %></td></tr>
<tr>
	<td height=14 class="titulor" align="center">Chapa</td>
	<td height=14 class="titulor" align="center">Nome</td>
	<td height=14 class="titulor" align="center">R.T.</td>
	<td height=14 class="titulor" align="center">Titulação</td>
	<td height=14 class="titulor" align="center">Admissão </td>
	<td height=14 class="titulor" align="center">Dia      </td>
	<td height=14 class="titulor" align="center">Turno    </td>
	<td height=14 class="titulor" align="center" colspan=2>Horário  </td>
	<td height=14 class="titulor" align="center">Turma    </td>
	<td height=14 class="titulor" align="center">Curso    </td>
	<td height=14 class="titulor" align="center">Disciplina</td>
	<td height=14 class="titulor" align="center">Aulas/CH </td>
	<td height=14 class="titulor" align="center" colspan=3>Observações</td>
</tr>
<%
linhas=linhas+2
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
	<td height=28 class="campor" ><%if lastchapa=rs("chapa") then response.write "" else response.write rs("chapa")%></td>
	<td height=28 class="campor" ><b><%if lastchapa=rs("chapa") then response.write "" else response.write rs("nome")%></td>
	<td height=28 class="campor" align="center"><%if lastchapa=rs("chapa") then response.write "" else response.write cargahoraria%></td>
	<td height=28 class="campor" ><%if lastchapa=rs("chapa") then response.write "" else response.write rs("instrucao")%></td>
	<td height=28 class="campor" align="center"><%if lastchapa=rs("chapa") then response.write "" else response.write rs("dataadmissao")%></td>
	<td height=28 class="campor" align="center"><%=diasem%></td>
	<td height=28 class="campor" align="center"><%=turno%></td>

	<td height=28 class="campor" nowrap align="center" width=6><%=rs("horini")%></td>
	<td height=28 class="campor" nowrap align="center" width=6><%=rs("horfim")%></td>
	
	<td height=28 class="campor" align="center" nowrap><%=rs("codtur")%></td>
	<td height=28 class="campor" ><%=rs("curso")%></td>
	<td height=28 class="campor" ><%=rs("materia")%></td>
	<td height=28 class="campor" align="center"><%=rs("taulas")%></td>
	<td height=28 class="campor" colspan=3><%=rs("portaria") & complemento%></td>
</tr>
<%
linhas=linhas+2
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
	<td height=14 class="titulor" colspan="12" align="left">Total <%=lastchapa %></td>
	<td height=14 class="campor" align="center"><%=tfaulas+tfacad+tfadm%></td>
	<td height=14 class="campor" align="center" nowrap>Grad: <b><%=tfaulas%></td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Adm: <b><%=tfadm%></td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Acad: <b><%=tfacad%></td>
</tr>
<tr><td class="campor" colspan="20"><hr></td></tr>
<tr>
	<td height=14 class="titulor" colspan="12">Total Geral</td>
	<td height=14 class="campor" align="center"><%=tgaulas+tgacad+tgadm%></font></td>
	<td height=14 class="campor" align="center" nowrap>Grad: <%=tgaulas%></font></td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Adm: <%=tgadm%></font></td>
	<td height=14 class="campor" align="center" nowrap>&nbsp;Acad: <%=tgacad%></font></td>
</tr>
<% linhas=linhas+2 %>
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
</body>
</html>
<%
conexao.close
set conexao=nothing
%>