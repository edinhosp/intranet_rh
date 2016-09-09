<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a12")="N" or session("a12")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Total da Carga Horária por Disciplina/Atribuição</title>
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
<!-- modelo do relatorio final -->

<!-- selecoes -->
<form method="POST" action="rel_ch_foto.asp" name="form">
<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=500>
<tr><td valign=top colspan=2>
<p style="margin-bottom: 0" class=realce><b>Seleções para o relatório &quot;Total da Carga Horária&quot;</b></p>
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
		"FROM grades_2 AS gr INNER JOIN g2cursoeve AS gc ON gr.coddoc=gc.coddoc " & _
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
"g.juntar, g.jturma, g.dividir, g.extra, g.demons, f.funcao " & _
"FROM g2ch g, dc_professor AS f, g2cursoeve gc, corporerm.dbo.umaterias m " & _
"WHERE g.chapa1=f.CHAPA collate database_default AND g.coddoc=gc.coddoc and m.codmat collate database_default=g.codmat and '" & dtaccess(database) & "' Between [inicio] And [termino] " & _
"group by g.coddoc, gc.CURSO, f.CODSECAO, f.NOME, g.chapa1, m.materia, f.DATAADMISSAO, f.CODSITUACAO, f.TITULACAOPAGA, f.INSTRUCAOMEC, g.turno, g.codtur, g.serie, g.turma, g.diasem, g.codmat, g.perlet, g.inicio, g.termino, g.juntar, g.jturma, g.dividir, g.extra, g.demons, funcao "
'"ORDER BY f.CODSECAO, f.NOME, g.chapa1, g.curso, g.materia; "
sql3="union all "
sql4="SELECT 2 AS tipoch, ni.coddoc, gc.curso, f.CODSECAO, f.NOME, ni.CHAPA, nn.NOMEACAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem, '' as a1,'' as a2, ch=case when ni.codeve is null or ni.codeve='' then 0 else ni.ch end, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), '' as diretor, ni.MAND_INI, ni.MAND_FIM, '" & dtaccess(database) & "' , ni.PORTARIA, ni.CARGO, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons, '' funcao " & _
"FROM n_indicacoes AS ni INNER JOIN dc_professor AS f ON ni.CHAPA=f.CHAPA collate database_default INNER JOIN n_nomeacoes AS nn ON ni.id_nomeacao=nn.id_nomeacao LEFT JOIN g2cursoeve gc ON ni.coddoc=gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [mand_ini] And [mand_fim] " 
'"ORDER BY f.CODSECAO, f.NOME, ni.CHAPA, ni.curso, nn.NOMEACAO; "
sql6="SELECT 3 AS tipoch, g.coddoc, gc.curso, f.CODSECAO, f.NOME, f.CHAPA, g.DESCRICAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem,'' as a1,'' as a2, g.CH, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), null as diretor, g.inicio, g.FIM, '" & dtaccess(database) & "', '' as portaria, '' as obs, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons, '' funcao " & _
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
		filtrow="WHERE sessao='" & session("usuariomaster") & "' "
		filtroh=""
		selecao="Seleção: todos registros"
	case "2" 'curso
		filtrow="WHERE sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING coddoc='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes com aulas/atividades no curso: " & request.form("cselecao")
	case "3" 'disciplina
		filtrow="WHERE sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING materia='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes com a disciplina: " & request.form("cselecao")
	case "4" 'professor
		filtrow="WHERE sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING ss.chapa='" & request.form("cselecao") & "' "
		selecao="Seleção: apenas o docente com a chapa: " & request.form("cselecao")
	case "5" 'setor
		filtrow="WHERE sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING codsecao='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes alocados na seção: " & request.form("cselecao")
	case "6" 'diretor
		filtrow="WHERE sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING diretor_='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes cujo Diretor do curso é " & request.form("cselecao")
	case "7" 'titulação
		filtrow="WHERE sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING titulacaopaga='" & request.form("cselecao") & "' "
		selecao="Seleção: docentes com a titulação: " & request.form("cselecao")
	case "8" 'carga horaria
		valor1=request.form("T1")
		valor2=request.form("T2")
		filtrow="WHERE sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING cargahoraria Between " & valor1 & " And " & valor2 & " "
		selecao="Seleção: docentes com carga horária total entre " & valor1 & " And " & valor2 & " horas "
	case "9" 'especial
		filtrow="WHERE sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING ss.chapa In (select chapa from zselecao where sessao='" & session("usuariomaster") & "') "
		selecao="Seleção: docentes específicos."
end select

sqla="SELECT CODSECAO, descricao as SECAO, NOME, ss.CHAPA, ss.tipoch, coddoc, curso, DATAADMISSAO, CODSITUACAO, titulacaopaga instrucao, instrucaomec, " & _
"materia, Sum(aulas) AS taulas, ch.cargahoraria, funcao " & _
"FROM (" & sql10 & ") as ss, ttcargahoraria_ch ch, corporerm.dbo.psecao s "
sqlb=Filtrow & " and ss.chapa=ch.chapa and ch.tipoch=4 and s.codigo=ss.codsecao and codsituacao in ('A','F','Z','E') "
sqlc="GROUP BY CODSECAO, descricao, NOME, ss.CHAPA, ss.tipoch, coddoc, curso, DATAADMISSAO, CODSITUACAO, titulacaopaga, instrucaomec, " & _
"materia, diretor_, PORTARIA, OBS, ch.cargahoraria, funcao "
sqld=Filtroh
sqle="ORDER BY NOME, ss.tipoch, curso, materia "

sql1=sqla & sqlb & sqlc & sqld & sqle
'response.write "<br>" & sql1
rs.Open sql1, ,adOpenStatic, adLockReadOnly

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
<p class=realce style="margin-top:0; margin-bottom:0">Total da Carga Horária por Disciplina/Atribuição em <%=database%></p>
<%
rs.movefirst:do while not rs.eof 
chapaatual=rs("chapa")
if lastchapa<>rs("chapa") then
%>
<table border="0" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo colspan=2 style="border-top:3px solid #000000" height="15">&nbsp;</td>
</tr>
<tr>
	<td class=campo width="590">
		<table border="0" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="100%">
		<tr>
			<td class=campo>Nome<br><b><%=rs("nome")%></td>
			<td class=campo>Titulação Paga:<br><b><%=rs("instrucaomec")%></td>
			<td class=campo>Titulação Real:<br><b><%=rs("instrucao")%></td>
		</tr>
		</table>
	</td>
	<td class=campo rowspan=30><img border="0" src="func_foto.asp?chapa=<%=rs("chapa")%>" width="100"></td>
</tr>
<tr>
	<td class=campo>
		<table border="0" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="100%">
		<tr>
			<td class=campo>Admissão<br><b><%=rs("dataadmissao")%><td>
			<td class=campo>Função:<br><b><%=rs("funcao")%></td>
			<td class=campo>Carga Horaria Total:<br><b><%=rs("cargahoraria")%></td>
		</tr>
		</table>
	</td>
</tr>	
<tr>
	<td class=campo>
		<table border="0" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="100%">
		<tr>
			<td class=titulo>Tipo</td>
			<td class=titulo>Curso</td>
			<td class=titulo>Materia/Atividade</td>
			<td class=titulo>Nº Aulas/Horas</td>
		</tr>
<%
end if 'last chapa
if rs("tipoch")=1 then tipo="Aulas"
if rs("tipoch")=2 then tipo="Atividades"
if rs("tipoch")=3 then tipo="Acadêmicas"
%>
		<tr>
			<td class=campo><%=tipo%></td>
			<td class=campo><%=rs("curso")%></td>
			<td class=campo><%=rs("materia")%></td>
			<td class=campo><%=rs("taulas")%></td>
		</tr>
<%
inicio=0
lastchapa=rs("chapa")
rs.movenext
if rs.eof then chapaatual="" else chapaatual=rs("chapa")
if chapaatual<>lastchapa then
%>
		</table>
	</td>
</tr>
</table>
<%
end if
loop
rs.close
set rs=nothing

else 'sem registros
end if 'recordcount

end if 'if do request.form

conexao.close
set conexao=nothing
%>
</body>
</html>