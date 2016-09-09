<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a93")="N" or session("a93")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Plano de Ensino - Relatório</title>
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
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("B1")="" then
%>
<!-- modelo do relatorio inicio -->
<!-- modelo do relatorio final -->

<!-- selecoes -->
<form method="POST" name="form" action="planorel.asp">
<table border=1 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=600>
<tr><td valign=top colspan=2>
	<p style="margin-bottom: 0" class=realce><b>Seleções para o relatório &quot;Plano de Ensino&quot;</b></p>
</td></tr>
<tr>
	<td class=titulor nowrap>Tipo da Seleção</td>
	<td class=titulor>Conteúdo da Seleção</td>
	</tr>
<tr>
	<td class="campot"r nowrap><select size="1" name="selecao" onChange="javascript:submit()">
		<option value="1" <%if request.form("selecao")="1" then response.write "selected"%> >Amostra</option>
		<option value="2" <%if request.form("selecao")="2" then response.write "selected"%> >Curso</option>
		<option value="3" <%if request.form("selecao")="3" then response.write "selected"%> >Disciplina</option>
		<option value="4" <%if request.form("selecao")="4" then response.write "selected"%> >Professor</option>
		<option value="5" <%if request.form("selecao")="5" then response.write "selected"%> >Aluno</option>
		</select>
	</td>
	<td class="campot"r>
<%
combo=0
select case request.form("selecao")
	case "2" 'curso
		combo=1:opcao=0:sqltemp="select p.coddoc + '/' + convert(nvarchar,p.grade) as codigo, c.curso + ' (Grade: ' + convert(nvarchar,p.grade) + ')' as descricao " & _
		"from grades_plano p, g2cursoeve c where p.coddoc=c.coddoc " & _
		"group by p.coddoc + '/' + convert(nvarchar,p.grade), c.curso + ' (Grade: ' + convert(nvarchar,p.grade) + ')',curso, grade order by curso, grade "
		sqltemp="select codpergrade codigo, descricao=curso+' - Grade ' + convert(nvarchar,descricao collate database_default) + ' ('+peri+' a '+perf+')' from grades_pe order by tpcurso, curso, peri "
	case "3" 'disciplina
		combo=1:opcao=0:sqltemp="SELECT m.materia as codigo, m.materia as descricao " & _
		"FROM grades_plano pe inner join corporerm.dbo.umaterias m on pe.codmat=m.codmat collate database_default " & _
		"GROUP BY m.materia " & _
		"ORDER BY m.materia"
	case "4" 'professor
		combo=1:opcao=1:sqltemp="SELECT g.chapa1 as codigo, f.nome as descricao " & _
		"FROM grades_plano pe, g2ch g, grades_aux_prof as f " & _
		"WHERE pe.codmat=g.codmat and g.chapa1=f.chapa " & _
		"GROUP BY g.chapa1, f.nome ORDER BY f.nome"
	case "5" 'aluno
		combo=2:opcao=1
end select
if combo=1 then
%>
<select size="1" name="cselecao" onChange="javascript:submit()">
	<option value="">Selecione...</option>
<%
rs.Open sqltemp, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
	<option value="<%=rs("codigo")%>" <%if request.form("cselecao")=rs("codigo") then response.write "selected"%>  ><%=rs("descricao")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<%
end if 'selecao combo 1

if combo=2 then
	response.write "<input type='text' name='cselecao' size='12' value='" & request.form("cselecao") & "' onchange='javascript:submit()'>"
end if 'selecao combo 2

if request.form("selecao")="5" and request.form("cselecao")<>"" then
	sql="select nome from corporerm.dbo.ealunos where matricula='" & request.form("cselecao") & "' "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then response.write " <font size=2><b>" & rs("nome") & "</b>"
	rs.close
end if

if request.form("selecao")="2" and request.form("cselecao")<>"" then
	codigo=request.form("cselecao")
	sql2="select distinct codcur, codper, grade from grades_pe where codpergrade='" & codigo & "' "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	sql="select distinct perlet from grades_plano where codcur='" & rs2("codcur") & "' and codper=" & rs2("codper") & " and grade=" & rs2("grade")
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	response.write "<select size='1' name='cperlet'>"
	if rs.recordcount>0 then
	rs.movefirst:do while not rs.eof 
	response.write "<option value='" & rs("perlet") & "'>" & rs("perlet") & "</option>"
	rs.movenext:loop
	end if
	rs.close
	rs2.close
	response.write "</select>"
end if

if request.form("selecao")="3" and request.form("cselecao")<>"" then
	codigo=request.form("cselecao")
	sql="select perlet from grades_plano p inner join corporerm.dbo.umaterias m on p.codmat=m.codmat collate database_default " & _
	"and m.materia='" & codigo & "' group by perlet order by perlet desc"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	response.write "<select size='1' name='cperlet'>"
	rs.movefirst:do while not rs.eof 
	response.write "<option value='" & rs("perlet") & "'>" & rs("perlet") & "</option>"
	rs.movenext:loop
	rs.close
	response.write "</select>"
end if

if request.form("selecao")="4" and request.form("cselecao")<>"" then
	codigo=request.form("cselecao")
	sql="select p.perlet from grades_plano p, g2ch g " & _
	"where g.codmat=p.codmat and g.coddoc=p.coddoc and g.codcur=p.codcur and g.codper=p.codper and g.grade=p.grade and g.serie=p.serie and g.perlet=p.perlet and g.chapa1='" & codigo & "' group by p.perlet order by p.perlet desc"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	response.write "<select size='1' name='cperlet'>"
	if rs.recordcount>0 then
		rs.movefirst:do while not rs.eof 
		response.write "<option value='" & rs("perlet") & "'>" & rs("perlet") & "</option>"
		rs.movenext:loop
	end if
	rs.close
	response.write "</select>"
end if
%>
	</td>
</tr>

<tr><td valign=top class="campot" width=10%>Campos a imprimir:
</td>
<td valign=top class="campolr">
	<input type="checkbox" name="cjust" value="ON" checked>Justificativa<br>
	<input type="checkbox" name="cemen" value="ON" checked>Ementa<br>
	<input type="checkbox" name="cobje" value="ON" checked>Objetivos Gerais<br>
	<input type="checkbox" name="cunid" value="ON" >Unidades Temáticas<br>
	<input type="checkbox" name="cmeto" value="ON" >Metodologia<br>
	<input type="checkbox" name="caval" value="ON" >Avaliação<br>
	<input type="checkbox" name="cbbas" value="ON" >Bibliografia Básica<br>
	<input type="checkbox" name="cbcom" value="ON" >Bibliografia Complementar<br>
</td></tr>


<tr><td valign=top colspan=2 class="campot">
	<p><input type="submit" class=button value="Visualizar Relatório" name="B1"></p>
</td></tr>

<tr><td valign=top class=campoe colspan=2 class="campot">
	<p style="margin-top: 0; margin-bottom: 0"><font color="#FF0000">Configure a página do seu navegador (Internet
	Explorer, Netscape, Mozilla, etc) no sentido RETRATO.</font></p>
</td></tr>


</table>

</form>
<%
end if  'if do request.form

'**********************************************************

if request.form("B1")<>"" then

filtro="":filtro2="":selecao=""
codigo=request.form("cselecao")
perlet=request.form("cperlet")

select case request.form("selecao")
	case "1" 'todos
		selecao="Seleção: todos registros"
		sql="SELECT top 5 p.id_plano, p.CODMAT, u.MATERIA, p.justificativa, p.ementa, p.objetivos_gerais, p.unidades_tematicas, p.metodologia, p.avaliacao, p.bibliografia, p.bibliografiac, " & _
		"p.CODdoc, c.CURSO, p.grade, p.serie, g.NAULASSEM, g.CARGAHORARIA, c.DEPTO, p.pa, p.perlet, p.codcur, p.codper, p.grade " & _
		"FROM (grades_plano p INNER JOIN corporerm.dbo.umaterias u ON p.CODMAT=u.CODMAT collate database_default) " & _
		"inner join corporerm.dbo.ugrade g on g.codcur=p.codcur and g.codper=p.codper and g.grade=p.grade and g.periodo=p.serie and g.codmat=u.codmat " & _
		"INNER JOIN g2cursoeve c ON p.CODdoc=c.coddoc " 
	case "2" 'curso
		sql2="select distinct codcur, codper, grade from grades_pe where codpergrade='" & codigo & "' "
		rs2.Open sql2, ,adOpenStatic, adLockReadOnly
		codcur=rs2("codcur"):codper=rs2("codper"):grade=rs2("grade")
		rs2.close
		selecao="Seleção: docentes com aulas/atividades no curso: " & request.form("cselecao")
		sql="SELECT p.id_plano, p.CODMAT, u.MATERIA, p.justificativa, p.ementa, p.objetivos_gerais, p.unidades_tematicas, p.metodologia, p.avaliacao, p.bibliografia, p.bibliografiac, " & _
		"p.CODdoc, c.CURSO, p.grade, p.serie, g.NAULASSEM, g.CARGAHORARIA, c.DEPTO, p.pa, p.perlet, p.codcur, p.codper, p.grade " & _
		"FROM (grades_plano p INNER JOIN corporerm.dbo.umaterias u ON p.CODMAT=u.CODMAT collate database_default) " & _
		"inner join corporerm.dbo.ugrade g on g.codcur=p.codcur and g.codper=p.codper and g.grade=p.grade and g.periodo=p.serie and g.codmat=u.codmat " & _
		"INNER JOIN g2cursoeve c ON p.CODdoc=c.coddoc " & _
		"WHERE p.codcur=" & codcur & " and p.codper=" & codper & " and p.grade=" & grade & " and perlet='" & perlet & "' " & _
		"order by p.serie, u.materia "
	case "3" 'disciplina
		selecao="Seleção: docentes com a disciplina: " & request.form("cselecao")
		sql="SELECT p.id_plano, p.CODMAT, u.MATERIA, p.justificativa, p.ementa, p.objetivos_gerais, p.unidades_tematicas, p.metodologia, p.avaliacao, p.bibliografia, p.bibliografiac, " & _
		"p.CODdoc, c.CURSO, p.grade, p.serie, g.NAULASSEM, g.CARGAHORARIA, c.DEPTO, p.pa, p.perlet, p.codcur, p.codper, p.grade " & _
		"FROM (grades_plano p INNER JOIN corporerm.dbo.umaterias u ON p.CODMAT=u.CODMAT collate database_default) " & _
		"inner join corporerm.dbo.ugrade g on g.codcur=p.codcur and g.codper=p.codper and g.grade=p.grade and g.periodo=p.serie and g.codmat=u.codmat " & _
		"INNER JOIN g2cursoeve c ON p.CODdoc=c.coddoc " & _
		"WHERE u.materia='" & codigo & "' and p.perlet='" & perlet & "' " & _
		"order by p.serie, u.materia "
	case "4" 'professor
		selecao="Seleção: apenas o docente com a chapa: " & request.form("cselecao")
		sql="SELECT p.id_plano, p.CODMAT, u.MATERIA, p.justificativa, p.ementa, p.objetivos_gerais, p.unidades_tematicas, p.metodologia, p.avaliacao, p.bibliografia, p.bibliografiac, " & _
		"p.CODdoc, c.CURSO, p.grade, p.serie, g.NAULASSEM, g.CARGAHORARIA, c.DEPTO, p.pa, p.perlet, p.codcur, p.codper, p.grade " & _
		"FROM (grades_plano p INNER JOIN corporerm.dbo.umaterias u ON p.CODMAT=u.CODMAT collate database_default) " & _
		"inner join corporerm.dbo.ugrade g on g.codcur=p.codcur and g.codper=p.codper and g.grade=p.grade and g.periodo=p.serie and g.codmat=u.codmat " & _
		"INNER JOIN g2cursoeve c ON p.CODdoc=c.coddoc " & _
		"inner join (select distinct perlet, chapa1, coddoc, codcur, codper, grade, codmat from g2ch) f on f.coddoc=p.coddoc and f.codcur=p.codcur and f.codper=p.codper and f.grade=p.grade and f.codmat=p.codmat and f.perlet=p.perlet " & _
		"WHERE f.chapa1='" & codigo & "' and p.perlet='" & perlet & "' " & _
		"order by p.perlet, u.materia "
	case "5" 'aluno
		selecao="Seleção: aluno: " & request.form("cselecao")
		sql="select p.id_plano, p.codmat, u.materia, a.mataluno, a.perletivo, a.codcur, a.codper, a.grade, a.codmat, a.codtur, p.serie, " & _
		"p.justificativa, p.ementa, p.objetivos_gerais, p.unidades_tematicas, p.metodologia, p.avaliacao, p.bibliografia, p.bibliografiac, pa, " & _
		"c.CODdoc, c.CURSO, g.NAULASSEM, g.CARGAHORARIA, c.DEPTO " & _
		"from corporerm.dbo.umatalun a inner join grades_plano p " & _
		"	on a.codcur=p.codcur and a.codper=p.codper and a.grade=p.grade and a.perletivo collate database_default=p.perlet and a.codmat collate database_default=p.codmat " & _
		"inner join corporerm.dbo.umaterias u on u.codmat=a.codmat " & _
		"inner join corporerm.dbo.ugrade g on g.codcur=a.codcur and g.codper=a.codper and g.grade=a.grade and g.codmat=a.codmat and g.periodo=p.serie " & _
		"inner join g2cursoeve c on p.coddoc=c.coddoc " & _
		"where a.mataluno='" & codigo & "' and a.status not in ('14','06') " & _
		"order by perletivo, u.materia "
end select

sqlz=sql:'response.write sql
rs.Open sqlz, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
inicio=1
'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'><tr>"
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
'rs.movenext:loop
'response.write "</table><p>"
'*************** fim teste **********************
sqlc="select c.nome, p.habilitacao from corporerm.dbo.ucursos c inner join corporerm.dbo.uperiodos p on p.codcur=c.codcur " & _
"where c.codcur=" & rs("codcur") & " and p.codper=" & rs("codper") & " "
rs2.Open sqlc, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	curso=rs2("nome"):habilitacao=rs2("habilitacao")
else
	curso="":habilitacao=""
end if
rs2.close

tlr="style='border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000'"
tl="style='border-top:1px solid #000000;border-left:1px solid #000000'"
tr="style='border-top:1px solid #000000;border-right:1px solid #000000'"
l="style='border-left:1px solid #000000'"
r="style='border-right:1px solid #000000'"
lr="style='border-left:1px solid #000000;border-right:1px solid #000000'"
blr="style='border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000'"
bl="style='border-bottom:1px solid #000000;border-left:1px solid #000000'"
br="style='border-bottom:1px solid #000000;border-right:1px solid #000000'"
b="style='border-bottom:1px solid #000000'"

dim texto(7),titulo(7),imprime(7):
do while not rs.eof
professores="":periodos=""
%>
<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td align="left" <%=tl%> width=110>
	<img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="110" alt=""></td>
	<td align="center" <%=tr%>>
	<b>PRÓ-REITORIA ACADÊMICA<br>PLANEJAMENTO ACADÊMICO</b></td>
</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo <%=l%>>Curso:</td>
	<td class=campo><b><%=curso & " / " & habilitacao %>
<b>
	</td>
	<td class=campo>Semestre/Período:</td>
	<td class=campo <%=r%>><b><%=rs("serie")%></td>
</tr>
<tr><td class=campo <%=l%>>Disciplina:</td>
	<td class=campo><b><%=rs("materia")%></td>
	<td class=campo>C/H Total:</td>
	<td class=campo <%=r%>><b><%=rs("cargahoraria")%></td>
</tr>
<tr><td class=campo <%=l%>>Professor:</td>
	<td class=campo width=290><b>
<%
if request.form("selecao")="5" then perlet=rs("perletivo") else perlet=request.form("cperlet")
sqlp="select g.chapa1, f.nome from g2ch g, corporerm.dbo.pfunc f where g.chapa1=f.chapa collate database_default and g.codmat='" & rs("codmat") & "' " & _
"and perlet='" & perlet & "' and coddoc='" & rs("coddoc") & "' group by g.chapa1, f.nome"
rsc.Open sqlp, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
	do while not rsc.eof
		professores=professores & rsc("nome")
		response.write rsc("nome")
		if rsc.recordcount>1 and rsc.absoluteposition<rsc.recordcount then response.write ", ":professores=professores & ", "
	rsc.movenext
	loop
end if
rsc.close
%>	
	</td>
	<td class=campo>C/H Semanal:</td>
	<td class=campo <%=r%>><b><%=rs("naulassem")%></td>
</tr>
<tr><td class=campo <%=bl%>>Departamento:</td>
	<td class=campo <%=b%>><b><%=rs("depto")%></td>
	<td class=campo <%=b%>>Turno:</td>
	<td class=campo <%=br%>><b>
<%
sqlp="select g.turno, t.tipo from g2ch g, eturnos t where g.codmat='" & rs("codmat") & "' and g.turno=t.codturno " & _
"and perlet='" & perlet & "' and coddoc='" & rs("coddoc") & "' group by g.turno, t.tipo "
rsc.Open sqlp, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
	do while not rsc.eof
		periodos=periodos & rsc("tipo")
		response.write rsc("tipo")
		if rsc.recordcount>1 and rsc.absoluteposition<rsc.recordcount then response.write "/":periodos=periodos & "/"
	rsc.movenext
	loop
end if
rsc.close
%>	
	</td>
</tr>
</table>
<%
tamanho=0
titulo(0)="JUSTIFICATIVA"             : texto(0)=rs("justificativa")     :if request.form("cjust")="ON" then imprime(0)=1
titulo(1)="EMENTA"                    : texto(1)=rs("ementa")            :if request.form("cemen")="ON" then imprime(1)=1
titulo(2)="OBJETIVOS GERAIS"          : texto(2)=rs("objetivos_gerais")  :if request.form("cobje")="ON" then imprime(2)=1
titulo(3)="UNIDADES TEMÁTICAS"        : texto(3)=rs("unidades_tematicas"):if request.form("cunid")="ON" then imprime(3)=1
titulo(4)="METODOLOGIA"               : texto(4)=rs("metodologia")       :if request.form("cmeto")="ON" then imprime(4)=1
titulo(5)="AVALIAÇÃO"                 : texto(5)=rs("avaliacao")         :if request.form("caval")="ON" then imprime(5)=1 
titulo(6)="BIBLIOGRAFIA BÁSICA"       : texto(6)=rs("bibliografia")      :if request.form("cbbas")="ON" then imprime(6)=1
titulo(7)="BIBLIOGRAFIA COMPLEMENTAR" : texto(7)=rs("bibliografiac")     :if request.form("cbcom")="ON" then imprime(7)=1
for a=0 to 7
	quadro=texto(a)
	if isnull(quadro)=false then quadro=replace(quadro,chr(13)&chr(10),"<br>")
	texto(a)=quadro
next

for a=0 to 7

if imprime(a)=1 then 

tam=len(texto(a)):tamanho=tamanho+tam
if tamanho>53*60 then
	response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td align="left" <%=tl%> width=110>
	<img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="110" alt=""></td>
	<td align="center" <%=tr%>>
	<b>PRÓ-REITORIA ACADÊMICA<br>PLANEJAMENTO ACADÊMICO</b></td>
</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo <%=l%>>Curso:</td>
	<td class=campo><b><%=rs("curso")%>
<b>
	</td>
	<td class=campo>Semestre/Período:</td>
	<td class=campo <%=r%>><b><%=rs("serie")%></td>
</tr>
<tr><td class=campo <%=l%>>Disciplina:</td>
	<td class=campo><b><%=rs("materia")%></td>
	<td class=campo>C/H Total:</td>
	<td class=campo <%=r%>><b><%=rs("cargahoraria")%></td>
</tr>
<tr><td class=campo <%=l%>>Professor:</td>
	<td class=campo width=290><b><%=professores%>	
	</td>
	<td class=campo>C/H Semanal:</td>
	<td class=campo <%=r%>><b><%=rs("naulassem")%></td>
</tr>
<tr><td class=campo <%=bl%>>Departamento:</td>
	<td class=campo <%=b%>><b><%=rs("depto")%></td>
	<td class=campo <%=b%>>Período:</td>
	<td class=campo <%=br%>><b><%=periodos%>
	</td>
</tr>
</table>
<%	
	tamanho=tam
end if
if a=6 then
	texto(6)=""
	sql="select p.id_biblio, p.id_plano, p.complementar, p.cod_acervo, p.ordem, p.referencia digitada, p.status, b.referencia pesquisada " & _
	"from grades_plano_biblio p left join pe_biblio b on b.cod_acervo=p.cod_acervo " & _
	"where id_plano=" & rs("id_plano") & " and complementar=0 " & _
	"order by ordem"
	rs2.Open sql, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
	if rs2("pesquisada")<>"" then referencia=rs2("pesquisada") else referencia=rs2("digitada")
	texto(6)=texto(6) & referencia & "<br>"
	rs2.movenext:loop
	rs2.close
end if
if a=7 then
	texto(7)=""
	sql="select p.id_biblio, p.id_plano, p.complementar, p.cod_acervo, p.ordem, p.referencia digitada, p.status, b.referencia pesquisada " & _
	"from grades_plano_biblio p left join pe_biblio b on b.cod_acervo=p.cod_acervo " & _
	"where id_plano=" & rs("id_plano") & " and complementar=1 " & _
	"order by ordem"
	rs2.Open sql, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
	if rs2("pesquisada")<>"" then referencia=rs2("pesquisada") else referencia=rs2("digitada")
	texto(7)=texto(7) & referencia & "<br>"
	rs2.movenext:loop
	rs2.close
end if

%>
<br><%'=tamanho & " / " & tamanho / 60%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo <%=tlr%> ><b><%=titulo(a)%></td></tr>
<tr><td class=campo <%=blr%> ><%=texto(a)%></td></tr>
</table>
<%
end if

next
%>

</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo align="left"><b><i>
<%
	if rs("pa")=1 or rs("pa")=true then valida="Validado pelo Planejamento Acadêmico." else valida="Não validado pelo Planejamento Acadêmico"
	response.write valida
%>
	</td>
	<td class="campop" align="right"><b><i>Período Letivo: <%=perlet%></td>
</tr>
</table>
<p style="margin-top:0;margin-bottom:0"><b><i>
<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página -->
rs.movenext
loop
rs.close
%>
</p>
<%
else 'sem registros
%>
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