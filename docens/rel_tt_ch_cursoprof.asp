<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
'revisado com divisao de codigos de cursos
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a13")="N" or session("a13")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Total da Carga Horária por Docente</title>
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
<table border=1 bordercolor=#000000 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=660>
<th>Modelo do relatório</th>
<tr><td valign=top>

<p class=realce style="margin-top:0; margin-bottom:0">Total da Carga Horária por Curso/Docente em 15/03/04</p>
<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td rowspan=2 class="titulor" align="center">Chapa</td>
	<td rowspan=2 class="titulor" align="center">Docente</td>
	<td rowspan=2 class="titulor" align="center">Titulação</td>
	<td rowspan=2 class="titulor" align="center">Admissão</td>
	<td colspan=4 class="titulor" align="center" style="border-right: 2 solid">Neste curso</td>
	<td colspan=4 class="titulor" align="center">No total</td>
</tr>
<tr>
	<td class="titulor" align="center">Aulas</td>
	<td class="titulor" align="center">Ativ.</td>
	<td class="titulor" align="center">Acad.</td>
	<td class="titulor" align="center" style="border-right: 2 solid">Total</td>
	<td class="titulor" align="center">Aulas</td>
	<td class="titulor" align="center">Ativ.</td>
	<td class="titulor" align="center">Acad.</td>
	<td class="titulor" align="center">Total</td>
</tr>

<tr><td class="grupo" colspan="12">&nbsp;Curso: FARMÁCIA</td></tr>

  <tr>
    <td class="campor">&nbsp;02312</td>
    <td class="campor">&nbsp;ANDREA MARIA GARRIDO DOS SANTOS</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;15/08/02</td>
    <td class="campor" align="center">5</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">5</td>
    <td class="campor" align="center">13</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">13</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;02134</td>
    <td class="campor">&nbsp;CASSIUS VINICIUS STEVANI</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;02/05/01</td>
    <td class="campor" align="center">3</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">3</td>
    <td class="campor" align="center">11</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">11</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;00622</td>
    <td class="campor">&nbsp;HILDA MACHADO DA SILVA LEITE</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;01/08/97</td>
    <td class="campor" align="center">3</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">3</td>
    <td class="campor" align="center">19</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">19</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;00422</td>
    <td class="campor">&nbsp;MARCIA HELENA BIAGGI ROSSI</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;01/08/95</td>
    <td class="campor" align="center">3</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">3</td>
    <td class="campor" align="center">19</td>
    <td class="campor" align="center">10</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">29</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;02357</td>
    <td class="campor">&nbsp;MARIA DE FATIMA PAREDES OLIVEIRA</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;03/02/03</td>
    <td class="campor" align="center">6</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">6</td>
    <td class="campor" align="center">18</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">18</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;02298</td>
    <td class="campor">&nbsp;MARLOS MEILUS</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;06/08/02</td>
    <td class="campor" align="center">3</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">3</td>
    <td class="campor" align="center">17</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">17</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;00670</td>
    <td class="campor">&nbsp;MERCEDES TOLEDO GRIJALBA</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;04/02/99</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">10</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">10</td>
    <td class="campor" align="center">20</td>
    <td class="campor" align="center">16</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">36</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;00609</td>
    <td class="campor">&nbsp;MIRIAM MITSUE HAYASHI</td>
    <td class="campor">&nbsp;Mestre</td>
    <td class="campor">&nbsp;07/02/97</td>
    <td class="campor" align="center">3</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">3</td>
    <td class="campor" align="center">17</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">17</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;00732</td>
    <td class="campor">&nbsp;PATRICIA SARTORELLI</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;01/08/00</td>
    <td class="campor" align="center">3</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">3</td>
    <td class="campor" align="center">11</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">11</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;02185</td>
    <td class="campor">&nbsp;PAULA CRISTINA VENEROSO</td>
    <td class="campor">&nbsp;Mestre</td>
    <td class="campor">&nbsp;07/08/01</td>
    <td class="campor" align="center">2</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">2</td>
    <td class="campor" align="center">16</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">16</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;02411</td>
    <td class="campor">&nbsp;PAULO BRANDAO</td>
    <td class="campor">&nbsp;Mestre</td>
    <td class="campor">&nbsp;01/08/03</td>
    <td class="campor" align="center">13</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">13</td>
    <td class="campor" align="center">17</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">17</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;00751</td>
    <td class="campor">&nbsp;PAULO ROBERTO DA CUNHA</td>
    <td class="campor">&nbsp;Mestre</td>
    <td class="campor">&nbsp;09/08/00</td>
    <td class="campor" align="center">2</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">2</td>
    <td class="campor" align="center">28</td>
    <td class="campor" align="center">2</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">30</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;02353</td>
    <td class="campor">&nbsp;REYNALDO MASCAGNI GATTI</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;03/02/03</td>
    <td class="campor" align="center">8</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">8</td>
    <td class="campor" align="center">16</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">16</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;02228</td>
    <td class="campor">&nbsp;RICARDO AUGUSTO ROTTER MONTIBELLER</td>
    <td class="campor">&nbsp;Mestre</td>
    <td class="campor">&nbsp;06/02/02</td>
    <td class="campor" align="center">3</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">3</td>
    <td class="campor" align="center">10</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">10</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;00433</td>
    <td class="campor">&nbsp;RICARDO LUIS DE SOUZA</td>
    <td class="campor">&nbsp;Mestre</td>
    <td class="campor">&nbsp;01/08/95</td>
    <td class="campor" align="center">4</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">4</td>
    <td class="campor" align="center">20</td>
    <td class="campor" align="center">1</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">21</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;02179</td>
    <td class="campor">&nbsp;RODRIGO ESTEVES DE LIMA LOPES</td>
    <td class="campor">&nbsp;Mestre</td>
    <td class="campor">&nbsp;01/08/01</td>
    <td class="campor" align="center">2</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">2</td>
    <td class="campor" align="center">20</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">20</td>
  </tr>

  <tr>
    <td class="campor">&nbsp;00692</td>
    <td class="campor">&nbsp;TELMA SUMIE MASUKO</td>
    <td class="campor">&nbsp;Doutor</td>
    <td class="campor">&nbsp;14/02/00</td>
    <td class="campor" align="center">3</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center" style="border-right: 2 solid">3</td>
    <td class="campor" align="center">6</td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center"></td>
    <td class="campor" align="center">6</td>
  </tr>

  <tr>
    <td class="titulor" colspan=4>&nbsp;Total FARMÁCIA</td>
    <td class="campor" align="center">66</td>
    <td class="campor" align="center">10</td>
    <td class="campor" align="center">0</td>
    <td class="campor" align="center" style="border-right: 2 solid">76</td>
    <td class="campor" colspan=4 align="center">&nbsp;</td>
  </tr>
  <tr>
    <td class="titulor" colspan=4>&nbsp;Total Geral</td>
    <td class="campor" align="center">66</td>
    <td class="campor" align="center">10</td>
    <td class="campor" align="center">0</td>
    <td class="campor" align="center" style="border-right: 2 solid">76</td>
    <td class="campor" colspan=4 align="center">&nbsp;</td>
  </tr>
</table>

</td></tr></table>
<!-- modelo do relatorio final -->

<!-- selecoes -->
<form method="POST" action="rel_tt_ch_cursoprof.asp" name="form">
<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=500>
<tr><td valign=top colspan=2>
<p style="margin-bottom: 0" class=realce><b>Seleções para o relatório &quot;Total da Carga Horária   por Curso/Docente&quot;</b></p>
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
response.write sqltemp
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
<tr><td class="campor" colspan=2>Sem quebra de página (imprime na sequência) <input type="checkbox" name="quebra" value="ON"></td></tr>
<tr><td class="campor" colspan=2>Quebra de página por curso <input type="checkbox" name="quebracurso" value="ON"></td></tr>
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
conexao.execute sql1

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
sql2="SELECT 1 AS tipoch, g.coddoc, gc.CURSO, f.CODSECAO, f.NOME, g.chapa1 chapa, m.materia collate database_default materia, f.DATAADMISSAO, f.CODSITUACAO, f.TITULACAOPAGA, f.INSTRUCAOMEC, instrucao, g.turno, g.codtur, g.serie, g.turma, g.diasem, " & _
"min(g.horini) horini, max(g.horfim) horfim, sum(g.ta) AS aulas, g.codmat, g.perlet, diretor_='', g.inicio, g.termino, '" & dtaccess(database) & "' AS [database], '' AS portaria, '' AS obs, " & _
"g.juntar, g.jturma, g.dividir, g.extra, g.demons " & _
"FROM g2ch g, dc_professor AS f, g2cursoeve gc, corporerm.dbo.umaterias m " & _
"WHERE g.chapa1=f.CHAPA collate database_default AND g.coddoc=gc.coddoc and m.codmat collate database_default=g.codmat and '" & dtaccess(database) & "' Between [inicio] And [termino] " & _
"group by g.coddoc, gc.CURSO, f.CODSECAO, f.NOME, g.chapa1, m.materia, f.DATAADMISSAO, f.CODSITUACAO, f.TITULACAOPAGA, f.INSTRUCAOMEC, instrucao, g.turno, g.codtur, g.serie, g.turma, g.diasem, g.codmat, g.perlet, g.inicio, g.termino, g.juntar, g.jturma, g.dividir, g.extra, g.demons "
'"ORDER BY f.CODSECAO, f.NOME, g.chapa1, g.curso, g.materia; "
sql3="union all "
sql4="SELECT 2 AS tipoch, ni.coddoc, gc.curso, f.CODSECAO, f.NOME, ni.CHAPA, nn.NOMEACAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, instrucao, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem, '' as a1,'' as a2, ch=case when ni.codeve is null or ni.codeve='' then 0 else ni.ch end, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), '' as diretor, ni.MAND_INI, ni.MAND_FIM, '" & dtaccess(database) & "' , ni.PORTARIA, ni.CARGO, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons " & _
"FROM n_indicacoes AS ni INNER JOIN dc_professor AS f ON ni.CHAPA=f.CHAPA collate database_default INNER JOIN n_nomeacoes AS nn ON ni.id_nomeacao=nn.id_nomeacao LEFT JOIN g2cursoeve gc ON ni.coddoc=gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [mand_ini] And [mand_fim] " 
'"ORDER BY f.CODSECAO, f.NOME, ni.CHAPA, ni.curso, nn.NOMEACAO; "
sql6="SELECT 3 AS tipoch, g.coddoc, gc.curso, f.CODSECAO, f.NOME, f.CHAPA, g.DESCRICAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, instrucao, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem,'' as a1,'' as a2, g.CH, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), null as diretor, g.inicio, g.FIM, '" & dtaccess(database) & "', '' as portaria, '' as obs, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons " & _
"FROM grades_rt AS g INNER JOIN dc_professor AS f ON g.CHAPA=f.CHAPA collate database_default LEFT JOIN g2cursoeve gc ON g.coddoc=gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [inicio] And [fim] " 
'"ORDER BY f.CODSECAO, f.NOME, f.CHAPA, g.curso, g.DESCRICAO; "
'response.write "<br>" & sql2 & "<br>"
'response.write "<br>" & sql4 & "<br>"
'response.write "<br>" & sql6 & "<br>"
sql10=sql2 & sql3 & sql4 & sql3 & sql6
'response.write "<br>" & sql10


select case request.form("selecao")
	case "1" 'todos
		filtrow="WHERE ch.sessao='" & session("usuariomaster") & "' "
		filtroh=""
		selecao="Seleção: todos registros"
	case "2" 'curso
		filtrow="WHERE ch.sessao='" & session("usuariomaster") & "' and coddoc='" & request.form("cselecao") & "' "
		filtroh=""
		selecao="Seleção: docentes com aulas/atividades no curso: " & request.form("cselecao")
	case "3" 'disciplina
		filtrow="WHERE ch.sessao='" & session("usuariomaster") & "' and materia='" & request.form("cselecao") & "' "
		filtroh=""
		selecao="Seleção: docentes com a disciplina: " & request.form("cselecao")
	case "4" 'professor
		filtrow="WHERE ch.sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING ss.chapa='" & request.form("cselecao") & "' "
		selecao="Seleção: apenas o docente com a chapa: " & request.form("cselecao")
	case "5" 'setor
		filtrow="WHERE ch.sessao='" & session("usuariomaster") & "' and codsecao='" & request.form("cselecao") & "' "
		filtroh=""
		selecao="Seleção: docentes alocados na seção: " & request.form("cselecao")
	case "6" 'diretor
		filtrow="WHERE ch.sessao='" & session("usuariomaster") & "' and diretor_='" & request.form("cselecao") & "' "
		filtroh=""
		selecao="Seleção: docentes cujo Diretor do curso é " & request.form("cselecao")
	case "7" 'titulação
		filtrow="WHERE ch.sessao='" & session("usuariomaster") & "' and titulacaopaga='" & request.form("cselecao") & "' "
		filtroh=""
		selecao="Seleção: docentes com a titulação: " & request.form("cselecao")
	case "8" 'carga horaria
		valor1=request.form("T1")
		valor2=request.form("T2")
		filtrow="WHERE ch.sessao='" & session("usuariomaster") & "' and [4] Between " & valor1 & " And " & valor2 & " "
		filtroh=""
		selecao="Seleção: docentes com carga horária total entre " & valor1 & " And " & valor2 & " horas "
	case "9" 'especial
		filtrow="WHERE ch.sessao='" & session("usuariomaster") & "' "
		filtroh="HAVING ss.chapa In (select chapa from zselecao where sessao='" & session("usuariomaster") & "') "
		selecao="Seleção: docentes específicos."
end select

sqla="SELECT NOME, ss.CHAPA, coddoc=case when coddoc is null or coddoc='' then '-' else coddoc end, curso, DATAADMISSAO, CODSITUACAO, titulacaopaga as GRAUINSTRUCAO, instrucao as INSTRUCAO, " & _
"sum(case when ss.tipoch=1 then aulas else 0 end) as ct1, sum(case when ss.tipoch=2 then aulas else 0 end) as ct2, " & _
"sum(case when ss.tipoch=3 then aulas else 0 end) as ct3, sum(aulas) as ct4, " & _
"min([1]) as st1, min([2]) as st2, min([3]) as st3, min([4]) as st4, count(ss.chapa + curso) as linhas " & _
"FROM (" & sql10 & ") as ss, ttcargahoraria_col ch " 
sqlb=Filtrow & " and ch.chapa=ss.chapa and ch.sessao='" & session("usuariomaster") & "' "
sqlc="GROUP BY case when coddoc is null or coddoc='' then '-' else coddoc end, curso, NOME, ss.CHAPA, DATAADMISSAO, CODSITUACAO, titulacaopaga, INSTRUCAO " & _
" "
sqld=Filtroh
sqle="ORDER BY case when coddoc is null or coddoc='' then '-' else coddoc end, NOME "

sql1=sqla & sqlb & sqlc & sqld & sqle
'response.write "<br/>" & sql1 & "<br/>"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
'response.write "<br>" & rs.recordcount

tfaulas=0:tfacad=0:tfadm=0:tfacad=0
tgaulas=0:tgacad=0:tgadm=0:tgacad=0
thorista=0:tparcial=0:tintegral=0
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
<p class=realce style="margin-top:0; margin-bottom:0">Total da Carga Horária por Curso/Docente em <%=database%></p>
<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td rowspan=2 class="titulor" align="center">Chapa</td>
	<td rowspan=2 class="titulor" align="center">Docente</td>
	<td rowspan=2 class="titulor" align="center">Titulação</td>
	<td rowspan=2 class="titulor" align="center">Admissão</td>
	<td colspan=4 class="titulor" align="center" style="border-right: 2 solid">Neste curso</td>
	<td colspan=4 class="titulor" align="center" style="border-right: 2 solid">No total</td>
	<td rowspan=2 class="titulor" align="center">Tipo<br>(atual)</td>
	<td rowspan=2 class="titulor" align="center">Regime</td>
	
</tr>
<tr>
	<td class="titulor" align="center">Aulas</td>
	<td class="titulor" align="center">Ativ.</td>
	<td class="titulor" align="center">Acad.</td>
	<td class="titulor" align="center" style="border-right: 2 solid">Total</td>
	<td class="titulor" align="center">Aulas</td>
	<td class="titulor" align="center">Ativ.</td>
	<td class="titulor" align="center">Acad.</td>
	<td class="titulor" align="center">Total</td>
</tr>
<%
linhas=2
tcaulas=0:tcativ=0:tcgeral=0:tcacad=0
tgaulas=0:tgativ=0:tggeral=0:tgacad=0
totcur=0:totger=0
rs.movefirst
do while not rs.eof 
chapach=rs("chapa")
session("chapa")=chapach
if inicio=1 then
%>
<tr><td class="grupo" colspan="14">&nbsp;Curso: <%=rs("curso")%></td></tr>
<%
linhas=linhas+1
end if

if request.form("quebra")="" then
if linhas>70 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p style='margin-top:0; margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<p class=realce style='margin-top:0; margin-bottom:0' class=""titulor"">Total da Carga Horária por Curso/Docente em " & database & "</p>"
	linhas=1
	response.write "<table border='1' bordercolor=#000000 cellpadding='1' cellspacing='0' style='border-collapse: collapse' width='650'>"
	response.write "<tr>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Chapa</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Docente</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Titulação</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Admissão</td>"
	response.write "<td colspan=4 class=""titulor"" align=""center"">Neste curso</td>"
	response.write "<td colspan=4 class=""titulor"" align=""center"">No total</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Tipo<br>(atual)</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Regime</td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td class=""titulor"" align=""center"">Aulas</td>"
	response.write "<td class=""titulor"" align=""center"">Ativ.</td>"
	response.write "<td class=""titulor"" align=""center"">Acad.</td>"
	response.write "<td class=""titulor"" align=""center"">Total</td>"
	response.write "<td class=""titulor"" align=""center"">Aulas</td>"
	response.write "<td class=""titulor"" align=""center"">Ativ.</td>"
	response.write "<td class=""titulor"" align=""center"">Acad.</td>"
	response.write "<td class=""titulor"" align=""center"">Total</td>"
	response.write "</tr>"
	linhas=linhas+1
end if
end if 'quebra

if inicio=0 then
	if lastcurso<>rs("coddoc") then
%>
  <tr>
    <td class="titulor" colspan=4>&nbsp;Total <%=lastcurso %> (<%=totcur%> prof.)</td>
    <td class="campor" align="center"><%=tcaulas%></td>
    <td class="campor" align="center"><%=tcativ%></td>
    <td class="campor" align="center"><%=tcacad%></td>
    <td class="campor" align="center" style="border-right: 2 solid"><%=tcgeral%></td>
    <td class="campor" colspan=4 align="center">&nbsp;</td>
    <td class="campor" colspan=2 align="center">Horista: <%=thorista%><br>Parcial: <%=tparcial%><br>Integral: <%=tintegral%> (<%=formatpercent((tintegral+tparcial)/totcur,0)%>)</td>
  </tr>
<%
	thorista=0:tparcial=0:tintegral=0
if request.form("quebracurso")="ON" then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p style='margin-top:0; margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<p class=realce style='margin-top:0; margin-bottom:0' class=""titulor"">Total da Carga Horária por Curso/Docente em " & database & "</p>"
	linhas=1
	response.write "<table border='1' bordercolor=#000000 cellpadding='1' cellspacing='0' style='border-collapse: collapse' width='650'>"
	response.write "<tr>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Chapa</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Docente</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Titulação</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Admissão</td>"
	response.write "<td colspan=4 class=""titulor"" align=""center"">Neste curso</td>"
	response.write "<td colspan=4 class=""titulor"" align=""center"">No total</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Tipo<br>(atual)</td>"
	response.write "<td rowspan=2 class=""titulor"" align=""center"">Regime</td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td class=""titulor"" align=""center"">Aulas</td>"
	response.write "<td class=""titulor"" align=""center"">Ativ.</td>"
	response.write "<td class=""titulor"" align=""center"">Acad.</td>"
	response.write "<td class=""titulor"" align=""center"">Total</td>"
	response.write "<td class=""titulor"" align=""center"">Aulas</td>"
	response.write "<td class=""titulor"" align=""center"">Ativ.</td>"
	response.write "<td class=""titulor"" align=""center"">Acad.</td>"
	response.write "<td class=""titulor"" align=""center"">Total</td>"
	response.write "</tr>"
	linhas=linhas+1
end if
%>
  <tr><td class="grupo" colspan=14>&nbsp;Curso: <%=rs("curso")%></td></tr>
<%
	linhas=linhas+2
	tcaulas=0:tcativ=0:tcgeral=0:tcacad=0:totcur=0:totcur=0
	end if
end if
%>
  <tr>
    <td class="campor">&nbsp;<%=rs("chapa")%></td>
    <td class="campor">&nbsp;<%=rs("nome")%></td>
    <td class="campor">&nbsp;<%=rs("instrucao")%></td>
    <td class="campor">&nbsp;<%=rs("dataadmissao")%></td>
    <td class="campor" align="center"><%if rs("ct1")<>0 then response.write rs("ct1")%></td>
    <td class="campor" align="center"><%if rs("ct2")<>0 then response.write rs("ct2")%></td>
    <td class="campor" align="center"><%if rs("ct3")<>0 then response.write rs("ct3")%></td>
    <td class="campor" align="center" style="border-right: 2 solid"><%if rs("ct4")<>0 then response.write rs("ct4")%></td>
    <td class="campor" align="center"><%if rs("st1")<>0 then response.write rs("st1")%></td>
    <td class="campor" align="center"><%if rs("st2")<>0 then response.write rs("st2")%></td>
    <td class="campor" align="center"><%if rs("st3")<>0 then response.write rs("st3")%></td>
    <td class="campor" align="center" style="border-right: 2 solid"><%if rs("st4")<>0 then response.write rs("st4")%></td>
<%
'************ rt/paga/rht
sqltipo="select tipo from quem_nomeacoes where chapa='" & rs("chapa") & "' "
rs2.Open sqltipo, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then tipo=rs2("tipo") else tipo="---"
rs2.close
st4=rs("st4"):st2=rs("st2")
if isnull(st4) then st4=0
if st4=0 and st2=0 then st4=1:st2=1
if st4>=40 then
	regime="Integral"
elseif st4>=12 and st4<40 and st2/st4>0.25 then
	regime="Parcial"
else
	regime="Horista"
end if

%>
    <td class="campor" align="center"><%=tipo%></td>
    <td class="campor" align="left"><%=regime%></td>
  </tr>
<%

linhas=linhas+1
inicio=0
lastcurso=rs("coddoc"):if lastcurso="" or isnull(lastcurso) or lastcurso=null then lastcurso="-"
lastchapa=rs("chapa")
if isnull(rs("ct1")) then var1=0 else var1=rs("ct1")
if isnull(rs("ct2")) then var2=0 else var2=rs("ct2")
if isnull(rs("ct3")) then var3=0 else var3=rs("ct3")
if isnull(rs("ct4")) then var4=0 else var4=rs("ct4")
tcaulas =tcaulas + var1
tcativ  =tcativ  + var2
tcacad  =tcacad  + var3
tcgeral =tcgeral + var4
tgaulas =tgaulas + var1
tgativ  =tgativ  + var2
tgacad  =tgacad  + var3
tggeral =tggeral + var4
totcur=totcur+1:totger=totger+1
if regime="Horista" then thorista=thorista+1
if regime="Parcial" then tparcial=tparcial+1
if regime="Integral" then tintegral=tintegral+1
rs.movenext
loop
rs.close
set rs=nothing
pconc=(tintegral+tparcial)/totcur
'response.write pconc
if pconc<0.2 then conceito=1
if pconc>=0.2 and pconc<0.33 then conceito=2
if pconc>=0.33 and pconc<0.6 then conceito=3
if pconc>=0.6 and pconc<0.8 then conceito=4
if pconc>=0.8 then conceito=5
%>
  <tr>
    <td class="titulor" colspan=4>&nbsp;Total <%=lastcurso %> (<%=totcur%> prof.)</td>
    <td class="campor" align="center"><%=tcaulas%></td>
    <td class="campor" align="center"><%=tcativ%></td>
    <td class="campor" align="center"><%=tcacad%></td>
    <td class="campor" align="center" style="border-right: 2 solid"><%=tcgeral%></td>
    <td class="campor" colspan=4 align="center" style="border-right: 2 solid">&nbsp;</td>
    <td class="campor" colspan=2 align="left">Horista: <%=thorista%><br>Parcial: <%=tparcial%><br>Integral: <%=tintegral%> (<%=formatpercent((tintegral+tparcial)/totcur,0)%>)</td>
  </tr>
  <tr>
    <td class="titulor" colspan=4>&nbsp;Total Geral</td>
    <td class="campor" align="center"><%=tgaulas%></td>
    <td class="campor" align="center"><%=tgativ%></td>
    <td class="campor" align="center"><%=tgacad%></td>
    <td class="campor" align="center" style="border-right: 2 solid"><%=tggeral%></td>
    <td class="campor" colspan=4 align="center" style="border-right: 2 solid">&nbsp;</td>
    <td class="campor" colspan=2 align="center">Conceito: <%=conceito%> </td>
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