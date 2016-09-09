<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a14")="N" or session("a14")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B1")="" then
	if request.form("agrupa")="" then agrupa=1 else agrupa=request.form("agrupa")
%>
<!-- modelo do relatorio inicio -->
<table border=1 bordercolor=#000000 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=660>
<th>Modelo do relatório</th>
<tr><td valign=top>
</td></tr></table>
<!-- modelo do relatorio final -->

<!-- selecoes -->
<form method="POST" name="form" action="rel_distri_grade.asp">
<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=500>
<tr><td valign=top colspan=2>
<p style="margin-bottom: 0" class=realce><b>Seleções para o relatório &quot;Distribuição pela Grade&quot;</b></p>
</td></tr>
	<tr>
		<td class="campor">Agrupa a cada <input type=text name=agrupa value="<%=agrupa%>" size=1 class=a> horas</td>
		<td class="campor">Data base para o relatório:&nbsp;<input type=text name=database value=<%=now%> size=10 class=a></td>
	</tr>
	<tr>
		<td class="titulor" nowrap>Tipo da Seleção</td>
		<td class="titulor">Conteúdo da Seleção</td>
	</tr>
	<tr>
		<td class="campor" nowrap><select size="1" name="selecao" onChange="javascript:submit()">
			<option value="1" <%if request.form("selecao")="1" then response.write "selected"%> >Todos</option>
			<option value="5" <%if request.form("selecao")="5" then response.write "selected"%> >Setor</option>
			<option value="7" <%if request.form("selecao")="7" then response.write "selected"%> >Titulação</option>
			<option value="8" <%if request.form("selecao")="8" then response.write "selected"%> >Carga Horária</option>
			<option value="9" <%if request.form("selecao")="9" then response.write "selected"%> >Especial</option>
		</select>
		</td>
		<td class="campor">
<%
combo=0
select case request.form("selecao")
	case "5" 'setor
		combo=1:sqltemp="select codsecao as codigo, secao as descricao from qry_funcionarios f, grades_chapa g where f.chapa=g.chapa collate database_default group by codsecao, secao "
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
%>

<%
if request.form("B1")<>"" then

filtro="":filtro2="":selecao=""
database=cdate(request.form("database"))
agrupa=request.form("agrupa")

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
sql2="SELECT 1 AS tipoch, g.coddoc, gc.CURSO, f.CODSECAO, f.NOME, g.chapa1 collate database_default as chapa, g.materia, f.DATAADMISSAO, f.CODSITUACAO, f.TITULACAOPAGA, " & _
"f.INSTRUCAOMEC, g.turno, g.codtur, g.serie, g.turma, g.diasem, g.a1, g.a2, g.a3, g.a4, g.a5, g.a6, g.ta AS aulas, g.codmat, gp.perlet, gp.perlet2, " & _
"diretor_=case when diretor is null and perlet3>='2004' and turno in (1,2,5) then 'Maria Celia Soares Hungria de Luca' else (case when diretor is null and perlet3>='2004' and turno in (3) then 'Luiz Carlos de Azevedo Filho' else diretor end) end, " & _
"g.inicio, g.termino, '" & dtaccess(database) & "' AS [database], '' AS portaria, '' AS obs, g.juntar, g.jturma, g.dividir, g.extra, g.demons " & _
"FROM ((g2ch AS g INNER JOIN dc_professor AS f ON g.chapa1 collate database_default=f.CHAPA collate database_default) INNER JOIN " & _
"(select coddoc, perlet, perlet2, enfase, pini, pfim, lanc, diretor from grades_per group by coddoc, perlet, perlet2, enfase, pini, pfim, lanc, diretor) " & _
"AS gp ON (gp.enfase = g.enfase) AND (g.perlet2 = gp.perlet2) AND (g.perlet = gp.perlet) AND (g.coddoc = gp.coddoc)) INNER JOIN g2cursoeve AS gc ON g.coddoc = gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [inicio] And [termino] and f.codsituacao in ('A','F','Z','E') "
'"ORDER BY f.CODSECAO, f.NOME, g.chapa1, g.curso, g.materia; "
sql3="union all "
sql4="SELECT 2 AS tipoch, ni.coddoc, gc.curso, f.CODSECAO, f.NOME, ni.CHAPA, nn.NOMEACAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem,'' as a1,'' as a2,'' as a3,'' as a4,'' as a5,'' as a6, ch=case when ni.codeve is null or ni.codeve='' then 0 else ni.ch end, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), cast(Year(getdate()) as char(4)), '' as diretor, ni.MAND_INI, ni.MAND_FIM, '" & dtaccess(database) & "' , ni.PORTARIA, ni.CARGO, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons " & _
"FROM n_indicacoes AS ni INNER JOIN dc_professor AS f ON ni.CHAPA=f.CHAPA collate database_default INNER JOIN n_nomeacoes AS nn ON ni.id_nomeacao=nn.id_nomeacao LEFT JOIN g2cursoeve gc ON ni.coddoc=gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [mand_ini] And [mand_fim] and f.codsituacao in ('A','F','Z','E') " 
'"ORDER BY f.CODSECAO, f.NOME, ni.CHAPA, ni.curso, nn.NOMEACAO; "
sql6="SELECT 3 AS tipoch, g.coddoc, gc.curso, f.CODSECAO, f.NOME, f.CHAPA, g.DESCRICAO, f.DATAADMISSAO, f.CODSITUACAO, f.titulacaopaga, " & _
"f.INSTRUCAOmec, 1 AS turno, '' as codtur, '' as serie, '' as turma, '' as diasem,'' as a1,'' as a2,'' as a3,'' as a4,'' as a5,'' as a6, g.CH, '' as codmat, " & _
"cast(Year(getdate()) as char(4)), cast(Year(getdate()) as char(4)), null as diretor, g.inicio, g.FIM, '" & dtaccess(database) & "', '' as portaria, '' as obs, 0 as juntar, '' as jturma, 0 as dividir, 0 as extra, 0 as demons " & _
"FROM grades_rt AS g INNER JOIN dc_professor AS f ON g.CHAPA=f.CHAPA collate database_default LEFT JOIN g2cursoeve gc ON g.coddoc=gc.coddoc " & _
"WHERE '" & dtaccess(database) & "' Between [inicio] And [fim] and f.codsituacao in ('A','F','Z','E') " 
'"ORDER BY f.CODSECAO, f.NOME, f.CHAPA, g.curso, g.DESCRICAO; "
'response.write "<br>" & sql2
'response.write "<br>" & sql4
'response.write "<br>" & sql6
sql10=sql2 & sql3 & sql4 & sql3 & sql6
'response.write "<br>" & sql10

select case request.form("selecao")
	case "1" 'todos
		filtrow="WHERE tipoch=4 and sessao='" & session("usuariomaster") & "' "
		filtroh=""
		selecao="Seleção: todos registros"
	case "5" 'setor
		filtrow="WHERE tipoch=4 and codsecao='" & request.form("cselecao") & "' and sessao='" & session("usuariomaster") & "' "
		filtroh=""
		selecao="Seleção: docentes alocados na seção: " & request.form("cselecao")
	case "7" 'titulação
		filtrow="WHERE tipoch=4 and titulacaopaga='" & request.form("cselecao") & "' and sessao='" & session("usuariomaster") & "' "
		filtroh=""
		selecao="Seleção: docentes com a titulação: " & request.form("cselecao")
	case "8" 'carga horaria
		valor1=request.form("T1")
		valor2=request.form("T2")
		filtrow="WHERE tipoch=4 and cargahoraria Between " & valor1 & " And " & valor2 & " and sessao='" & session("usuariomaster") & "' "
		filtroh=""
		selecao="Seleção: docentes com carga horária total entre " & valor1 & " And " & valor2 & " horas "
	case "9" 'especial
		filtrow="WHERE tipoch=4 and sessao='" & session("usuariomaster") & "' "
		filtrow="WHERE tipoch=4 and sessao='" & session("usuariomaster") & "' and t.chapa In (select chapa from zselecao where sessao='" & session("usuariomaster") & "') "
		filtroh=""
		selecao="Seleção: docentes específicos."
end select

sqla="SELECT Partition([cargahoraria],1,1000," & agrupa & ") AS Particao, Val(Partition([cargahoraria],1,1000," & agrupa & ")) AS Numero, " & _
"Count(t.CHAPA) AS Freq " & _
"FROM ttcargahoraria_ch t INNER JOIN dc_professor f ON t.CHAPA=f.CHAPA " 
sqlb=Filtrow 
sqlc="GROUP BY Partition([cargahoraria],1,1000," & agrupa & "), Val(Partition([cargahoraria],1,1000," & agrupa & ")) "
sqld=Filtroh
sqle="ORDER BY Partition([cargahoraria],1,1000," & agrupa & ") "
sql1=sqla & sqlb & sqlc & sqld & sqle

sql1="select particao, particao as numero, count(chapa) as freq from (" & _
"select round(cargahoraria/" & agrupa & "+0.49,0)*" & agrupa & " as particao, cargahoraria, t.chapa " & _
"FROM ttcargahoraria_ch t INNER JOIN dc_professor f ON t.CHAPA=f.CHAPA collate database_default " & filtrow & " and f.codsituacao in ('A','F','Z','E') " & _
") t " & _
"group by particao "

rs.Open sql1', ,adOpenStatic, adLockReadOnly

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
	
<p class=titulo>Distribuição de docentes pela grade de horas
<table border="1" cellpadding="2" cellspacing="1" style="border-collapse: collapse" width="500">
  <tr>
    <td class=titulo align="center">Grade de Horas</td>
    <td class=titulo align="center">Quantidade de Docentes</td>
  </tr>
<%
total=0
linhas=2
rs.movefirst
do while not rs.eof 
temp=trim(rs("particao"))
if left(temp,1)=":" then
	temp=replace(temp,":","")
else
	temp=replace(temp,": "," e ")
end if

%>
  <tr>
    <td class=campo>&nbsp;entre <%=temp%> horas aulas semanais</td>
    <td class=campo align="center">&nbsp;<%=rs("freq")%> docentes</td>
  </tr>
<%
linhas=linhas+1
total=total+rs("freq")
rs.movenext
loop
%>
  <tr>
    <td class=titulo>&nbsp;</td>
    <td class=campo align="center">&nbsp;<%=total %> docentes</td>
  </tr>
</table>
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
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>