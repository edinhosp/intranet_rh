<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a11")="N" or session("a11")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Carga horária acima do limite</title>
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach
dim rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
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

<p class=realce style="margin-top:0; margin-bottom:0">Carga Horária acima do limite em 15/03/04</p>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campor">
	<p align="right"><font size="1">Regras da seleção<br>
	Total de aulas &gt; 20 horas semanais<br>
	Total geral &gt; 40 horas semanais (aulas e atividades)</font></td>
</tr>
</table>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="titulor" align="center">Chapa</td>
	<td class="titulor" align="center">Nome</td>
	<td class="titulor" align="center">Admissão</td>
	<td class="titulor" align="center">Curso/Seção</td>
	<td class="titulor" align="center">Grad.</td>
	<td class="titulor" align="center">Ativ.</td>
	<td class="titulor" align="center">Acad.</td>
	<td class="titulor" align="center">T.Geral</td>
</tr>

<tr>
	<td class="campor" align="center">01080</td>
	<td class="campor">ALEXANDRE MARCOS DE MATTOS PIRES FERREIRA</td>
	<td class="campor" align="center">09/02/02</td>
	<td class="campor">CURSO DE DIREITO</td>
	<td class="campor" align="center">22</td>
	<td class="campor" align="center">0</td>
	<td class="campor" align="center">0</td>
	<td class="campor" align="center">22</td>
</tr>

<tr>
	<td class="campor" align="center">00755</td>
	<td class="campor">CARLOS ROBERTO SALIMENO</td>
	<td class="campor" align="center">07/02/01</td>
	<td class="campor">CURSO DE ADMINISTRACAO DE EMPRESAS</td>
	<td class="campor" align="center">24</td>
	<td class="campor" align="center">0</td>
	<td class="campor" align="center">0</td>
	<td class="campor" align="center">24</td>
</tr>

<tr>
	<td class="campor" align="center">00754</td>
	<td class="campor">EDUARDO RODRIGUES DA CUNHA GUASCO</td>
	<td class="campor" align="center">10/08/00</td>
	<td class="campor">CURSO DE ADMINISTRACAO DE EMPRESAS</td>
	<td class="campor" align="center">24</td>
	<td class="campor" align="center">0</td>
	<td class="campor" align="center">0</td>
	<td class="campor" align="center">24</td>
</tr>

</table>

</td></tr></table>
<!-- modelo do relatorio final -->

<form method="POST" action="rel_ch_limite.asp" name="form">
<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=500>
<tr><td valign=top colspan=2>
<p style="margin-bottom: 0" class=realce><b>Seleções para o relatório &quot;Carga Horária acima do limite&quot;</b></p>
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
			<option value="6" <%if request.form("selecao")="6" then response.write "selected"%> >Diretor</option>
			<option value="7" <%if request.form("selecao")="7" then response.write "selected"%> >Titulação</option>
			<option value="9" <%if request.form("selecao")="9" then response.write "selected"%> >Especial</option>
		</select>
		</td>
		<td class="campor">
<%
combo=0
select case request.form("selecao")
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

'***************************************************************

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
		filtrow="WHERE (ch.sessao='" & session("usuariomaster") & "') or (ch.sessao='" & session("usuariomaster") & "') "
		filtroh="HAVING (max(case when ch.tipoch=1 then cargahoraria else 0 end)>20) or (max(case when ch.tipoch=4 then cargahoraria else 0 end)>40) "
		selecao="Seleção: todos registros"
	case "6" 'diretor
		filtrow="WHERE (diretor_='" & request.form("cselecao") & "' and ch.sessao='" & session("usuariomaster") & "') " & _
		"           OR (diretor_='" & request.form("cselecao") & "' and ch.sessao='" & session("usuariomaster") & "')"
		filtroh="HAVING (max(case when ch.tipoch=1 then cargahoraria else 0 end)>20) or (max(case when ch.tipoch=4 then cargahoraria else 0 end)>40) "
		selecao="Seleção: docentes cujo Diretor do curso é " & request.form("cselecao")
	case "7" 'titulação
		filtrow="WHERE (titulacaopaga='" & request.form("cselecao") & "' and ch.sessao='" & session("usuariomaster") & "') " & _
		"           OR (titulacaopaga='" & request.form("cselecao") & "' and ch.sessao='" & session("usuariomaster") & "') "
		filtroh="HAVING (max(case when ch.tipoch=1 then cargahoraria else 0 end)>20) or (max(case when ch.tipoch=4 then cargahoraria else 0 end)>40) "
		selecao="Seleção: docentes com a titulação: " & request.form("cselecao")
	case "9" 'especial
		filtrow="WHERE (ss.chapa In (select chapa from zselecao where sessao='" & session("usuariomaster") & "') and ch.sessao='" & session("usuariomaster") & "') " & _
		"           OR (ss.chapa In (select chapa from zselecao where sessao='" & session("usuariomaster") & "') and ch.sessao='" & session("usuariomaster") & "') "
		filtroh="HAVING (max(case when ch.tipoch=1 then cargahoraria else 0 end)>20) or (max(case when ch.tipoch=4 then cargahoraria else 0 end)>40) "
		selecao="Seleção: docentes específicos."
end select
'"max(iif(ch.tipoch=1,cargahoraria,0)) as t1, max(iif(ch.tipoch=2,cargahoraria,0)) as t2, " & _
'"max(iif(ch.tipoch=3,cargahoraria,0)) as t3, max(iif(ch.tipoch=4,cargahoraria,0)) as t4 " & _

sqla="SELECT CODSECAO, descricao as secao, NOME, ss.CHAPA, DATAADMISSAO, CODSITUACAO, titulacaopaga as GRAUINSTRUCAO, instrucaomec as INSTRUCAO, " & _
"max(case when ch.tipoch=1 then cargahoraria else 0 end) as t1, max(case when ch.tipoch=2 then cargahoraria else 0 end) as t2, " & _
"max(case when ch.tipoch=3 then cargahoraria else 0 end) as t3, max(case when ch.tipoch=4 then cargahoraria else 0 end) as t4 " & _
"FROM ( (" & sql10 & ") as ss INNER JOIN ttcargahoraria_ch AS ch ON ss.chapa=ch.CHAPA) INNER JOIN corporerm.dbo.PSECAO s ON ss.CODSECAO=s.CODIGO "
sqlb=Filtrow
sqlc="GROUP BY CODSECAO, descricao, NOME, ss.CHAPA, DATAADMISSAO, CODSITUACAO, titulacaopaga, INSTRUCAOmec " 
sqld=Filtroh
sqle="ORDER BY CODSECAO, descricao, NOME "

sqlz=sqla & sqlb & sqlc & sqld & sqle
rs.Open sqlz, ,adOpenStatic, adLockReadOnly

%>
<p class=realce style="margin-top:0; margin-bottom:0">Carga Horária acima do limite em <%=database%></p>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor">
	<p align="right"><font size="1">Regras da seleção<br>
	Total de aulas &gt; 20 horas semanais<br>
	Total geral &gt; 40 horas semanais (aulas e atividades)</font></td>
</tr>
</table>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="titulor" align="center">Chapa</td>
	<td class="titulor" align="center">Nome</td>
	<td class="titulor" align="center">Admissão</td>
	<td class="titulor" align="center">Curso/Seção</td>
	<td class="titulor" align="center">Grad.</td>
	<td class="titulor" align="center">Ativ.</td>
	<td class="titulor" align="center">Acad.</td>
	<td class="titulor" align="center">T.Geral</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
chapach=rs("chapa")
session("chapa")=chapach
%>
<tr>
	<td class="campor" align="center">
		<a class=r href="docente_ver.asp?chapa=<%=rs("chapa")%>&nome=<%=rs("nome")%>" onclick="NewWindow(this.href,'CadastroProfessor','645','480','yes','center');return false" onfocus="this.blur()">	
		<%=rs("chapa")%></a></td>
	<td class="campor"><%=rs("nome") %></td>
	<td class="campor" align="center"><%=rs("dataadmissao")%></td>
	<td class="campor"><%=rs("secao")%></td>
	<td class="campor" align="center"><%=rs("t1")%></td>
	<td class="campor" align="center"><%=rs("t2")%></td>
	<td class="campor" align="center"><%=rs("t3")%></td>
	<td class="campor" align="center"><%=rs("t4")%></td>
</tr>
<%
rs.movenext
loop
end if 'rs.recordcount
rs.close
set rs=nothing
%>
</table>
<p><i><font size="1" color="#0000FF"><b><%=selecao %></b></font></i></p>
<%	pagina=pagina+1
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
%>
<%
end if 'if do request.form
%>
<%
conexao.close
set conexao=nothing
%>
</body>
</html>