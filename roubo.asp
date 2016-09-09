<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Engenharia de Computação - 1C</TITLE>
<link rel="stylesheet" type="text/css" href="diversos.css">
</HEAD>
<BODY>
<%
dim conexao,rs,marc(6), formato(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
set rs4=server.createobject ("ADODB.Recordset")
Set rs4.ActiveConnection = conexao
sql="select top 100 chapa1 as chapa from g2ch where codcur=41 and perlet3='20051' and serie=1 and turma='C' group by chapa1 order by chapa1 "
rs3.Open sql, ,adOpenStatic, adLockReadOnly
rs3.movefirst
do while not rs3.eof

sqla="select f.nome, f.codsituacao, f.chapa, f.dataadmissao, c.nome as funcao, f.codsecao, s.descricao as secao, f.estadocivil, " & _
"f.grauinstrucao, f.rua, f.numero, f.complemento, f.bairro, f.cidade, f.cep, f.telefone1, f.telefone2, f.telefone3, " & _
"f.fax, f.sexo, f.dtnascimento, f.email, f.cpf, f.cartidentidade, f.ufcartident, f.pispasep, f.codpessoa, f.datademissao, " & _
"ss.descricao as sit, i.descricao as titulacao, f.dtemissaoident, f.mae " & _
"from dc_professor f, pfuncao c, psecao s, pcodsituacao ss, pcodinstrucao i " & _
"where f.codfuncao = c.codigo and f.codsecao = s.codigo " & _
"and f.codsituacao = ss.codcliente and f.grauinstrucao=i.codcliente "
sqlb="and f.CHAPA='" & rs3("chapa") & "' "
sqlc="ORDER BY f.CHAPA"
sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if request.form("datad")<>"" then datad=formatdatetime(request.form("datad"),2) else datad=formatdatetime(now(),2)
if request.form("datad")="" and rs("codsituacao")="D" then datad=rs("datademissao")
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>CADASTRO DE DOCENTES</p>
<input type="hidden" name="chapa" value="<%=rs("chapa")%>">
<input type="hidden" name="nome" value="<%=rs("nome")%>">
<%
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
tabela=615
tbfoto=150
%>

<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
<tr>
	<td valign="top" class=campo>
<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
<tr>
	<td class=titulo>Nome:    </td>
	<td class=titulo>Situação:</td>
	<td class=titulo>Chapa:   </td>
</tr>
<tr>
	<td class=campo><b><%=rs("nome")%></b></td>
	<td class=campo><%=rs("sit")%>&nbsp;<%if rs("codsituacao")="D" then response.write rs("datademissao")%></td>
	<td class=campo><%=rs("chapa")%>      </td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width=<%=tabela-tbfoto%>>
<tr><td class=grupo>Dados Acadêmicos</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
<tr>
	<td class=titulo>Admissão:</td>
	<td class=titulo>Função:  </td>
	<td class=titulo>Seção:   </td>
	<td class=titulo>Instrução/Titulação:</td>
</tr>
<tr>
	<td class=campo><%=rs("dataadmissao")%></td>
	<td class=campo><%=rs("funcao")%>      </td>
	<td class=campo><%=rs("codsecao")%>&nbsp;<%=rs("secao")%></td>
	<td class=campo><%=rs("titulacao")%>    </td>
</tr>
</table>

    </td>
	<td width="<%=tbfoto%>" valign="top">
		<img border="0" src="func_foto.asp?chapa=<%=rs("chapa")%>"  width="<%=tbfoto%>">
	</td>
</tr>
</table>

<!-- grades horaria de aulas -->
<%
'rs.movenext:loop

sql2="select * from g2ch " & _
"where deletada=0 and chapa1='" & rs("chapa") & "' " & _
"and #" & dtaccess(datad) & "# between inicio and termino " & _
"order by curso, diasem, turno "
sql2="select perlet,codcur,curso,turno,serie,turma,codtur,diasem,sum(aula) as aulas,codmat,materia,chapa1,inicio,termino,deletada," & _
"ativo,juntar,jturma,dividir,dturma,extra,demons,prof,min(horini) as iaula,max(horfim) as faula from g2ch " & _
"where deletada=0 and chapa1='" & rs("chapa") & "' " & _
"and #" & dtaccess(datad) & "# between inicio and termino " & _
"group by perlet,codcur,curso,turno,serie,turma,codtur,diasem,codmat,materia,chapa1,inicio,termino,deletada," & _
"ativo,juntar,jturma,dividir,dturma,extra,demons,prof " & _
"order by curso, diasem, turno "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
cortitle="#FFFFFF"
bgtitle="#0000FF"
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr><td class=grupo colspan=19>Grade de atribuição de aulas em 
	<input type="text" name="datad" size="12" maxlength="12" value="<%=datad%>" class=form_apt onChange="javascript:submit()">
	</td></tr>
<%
taulas=0
if rs2.recordcount>0 then
%>
<tr>
	<td class=titulo align="center">#</td>
	<td class=titulor align="center">P.Let.</td>
	<td class=titulo align="center">Curso/Disciplina</td>
	<td class=titulor align="center">Tur</td>
	<td class=titulor align="center">Dia</td>
	<td class=fundor align="center">Turno</td>
	<td class=titulor align="center" colspan=2 width=60>Horario</td>
	<td class=titulo align="center">Nr.</td>
	<td class=fundor align="center">Alunos</td>
	<td class=titulor align="center">Junta</td>
	<td class=titulor align="center">Divide</td>
	<td class=titulor align="center">Extra</td>
	<td class=fundor align="center">Consid.</td>
	<td class=titulo align="center">Período</td>
</tr>
<%
rs2.movefirst
do while not rs2.eof 
'aulas=rs2("ta")
classe="campor"
if rs2("codcur")="41" and rs2("serie")="1" and rs2("turma")="C" then classe="fundor" else classe="campor"
'if rs2("a1")=1 then classe1="fundor" else classe1="campor"
'if rs2("a2")=1 then classe2="fundor" else classe2="campor"
'if rs2("a3")=1 then classe3="fundor" else classe3="campor"
'if rs2("a4")=1 then classe4="fundor" else classe4="campor"
'if rs2("a5")=1 then classe5="fundor" else classe5="campor"
'if rs2("a6")=1 then classe6="fundor" else classe6="campor"

if lastcurso<>rs2("curso") then
	response.write "<tr>"
	response.write "<td class="campol" colspan=15><b>" & rs2("curso") & "</td>"
	response.write "</tr>"
end if
if rs2("turno")="1" then turno="Mat"
if rs2("turno")="2" then turno="Vesp"
if rs2("turno")="3" then turno="Not"
if rs2("turno")="5" then turno="Vesp-EF"
if rs2("turno")="61" then turno="Int.M"
if rs2("turno")="62" then turno="Int.V"
if rs2("turno")="6" then turno="Int."

sql9="SELECT Count(UMATRICPL.MATALUNO) AS alunos " & _
"FROM USITMAT AS USITMAT_1 INNER JOIN ((UMATRICPL INNER JOIN UMATALUN ON (UMATRICPL.CODFILIAL = UMATALUN.CODFILIAL) AND " & _
"(UMATRICPL.CODCOLIGADA = UMATALUN.CODCOLIGADA) AND (UMATRICPL.MATALUNO = UMATALUN.MATALUNO) AND " & _
"(UMATRICPL.PERLETIVO = UMATALUN.PERLETIVO) AND (UMATRICPL.CODCUR = UMATALUN.CODCUR) AND (UMATRICPL.CODPER = UMATALUN.CODPER) " & _
"AND (UMATRICPL.GRADE = UMATALUN.GRADE)) INNER JOIN UMATERIAS ON (UMATALUN.CODMAT = UMATERIAS.CODMAT) AND " & _
"(UMATALUN.CODCOLIGADA = UMATERIAS.CODCOLIGADA)) ON USITMAT_1.CODSITMAT = UMATALUN.STATUS " & _
"WHERE (((UMATALUN.STATUS) In ('01','07','08','09','10','15','18','19','20','47','48','46')) " & _
"AND ((UMATALUN.CODTUR)='" & rs2("codtur") & "') " & _
"AND ((UMATRICPL.PERLETIVO)='" & rs2("perlet") & "') " & _
"AND ((UMATRICPL.CODCUR)=" & rs2("codcur")& ") " & _
"AND ((UMATALUN.CODMAT)='" & rs2("codmat") & "')) "
rs4.Open sql9, ,adOpenStatic, adLockReadOnly
if rs4.recordcount>0 then alunos=rs4("alunos") else alunos="-"
rs4.close
%>
<tr>
	<td class=<%=classe%> align="center"><%=rs2.absoluteposition%></td>
	<td class=<%=classe%> align="center"><%=rs2("perlet")%></td>
	<td class=<%=classe%> ><font color=blue><%=rs2("materia")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=rs2("serie") & rs2("turma")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=weekdayname(rs2("diasem"),1)%></td>
	<td class=<%=classe%> align="center"><%=turno%></td>
	
	<td class=<%=classe%> align="center" width=10><%=rs2("iaula")%></td>
	<td class=<%=classe%> align="center" width=10><%=rs2("faula")%></td>
	
	<td class=<%=classe%> align="center">&nbsp;<%=rs2("aulas")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=alunos%></td>

	<td class=<%=classe%> align="center">&nbsp;<%if rs2("juntar")=true then response.write "<font face='Wingdings'>ü</font>" & rs2("jturma") %></td>
	<td class=<%=classe%> align="center">&nbsp;<%if rs2("dividir")=true then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center">&nbsp;<%if rs2("extra")=true then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center">&nbsp;<%if rs2("demons")=false then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center" nowrap><%=rs2("inicio") & " a " & rs2("termino")%></td>
</tr>
<%
taulas=taulas+rs2("aulas")
lastcurso=rs2("curso")
rs2.movenext
loop
%>
<tr>
	<td class="campoa" colspan=8><b>Total de horas semanais em docência</td>
	<td class="campot" align="center"><%=taulas%></td>
	<td class="campoa" colspan=6></td>
</tr>
<%
end if 'recordcount rs2
rs2.close
%>
<!-- </table> -->

<!-- grades atribuicoes/nomeacoes -->
<%
sql2="SELECT i.id_nomeacao, n.NOMEACAO, i.id_indicado, i.CHAPA, i.NOME, i.PORTARIA, " & _
"i.codeve, i.CARGO, i.MAND_INI, i.MAND_FIM, i.CH, i.obs, i.contrato " & _
"FROM n_indicacoes as i, n_nomeacoes as n " & _
"WHERE i.id_nomeacao = n.id_nomeacao and i.chapa='" & rs("chapa") & "' " & _
"and #" & dtaccess(datad) & "# between mand_ini and mand_fim " & _
"ORDER BY n.nomeacao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<!-- <table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>"> -->
<%
tadm=0
if rs2.recordcount>0 then
%>
<tr><td class=grupo colspan=15>Nomeações e Atividades</td></tr>
<tr>
	<td class=titulo align="center">#</td>
	<td class=titulo align="center" colspan=2>Nomeação</td>
	<td class=titulo align="center" colspan=5>Curso/cargo/obs.</td>
	<td class=titulo align="center">Nr.</td>
	<td class=titulo align="center" colspan=4>Portaria</td>
	<td class=titulor align="center">Folha</td>
	<td class=titulo align="center">Período</td>
</tr>
<%
rs2.movefirst
do while not rs2.eof 
classe="campor"
%>
<tr>
	<td class=<%=classe%> align="center"><%=rs2.absoluteposition%></td>
	<td class=<%=classe%> colspan=2><%=(rs2("nomeacao"))%></td>
	<td class=<%=classe%> colspan=5><font color=blue><%=(rs2("cargo"))%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=rs2("ch")%></td>
	<td class=<%=classe%> colspan=4><%=rs2("portaria")%></td>
	<td class=<%=classe%> align="center"><%if rs2("codeve")<>"" then response.write rs2("codeve") else response.write "<font color=brown>S/ ônus"%></td>
	<td class=<%=classe%> align="center"><%=rs2("mand_ini") & " a " & rs2("mand_fim")%></td>
</tr>
<%
if rs2("ch")<>"" and rs2("codeve")<>"" then tadm=tadm+cdbl(rs2("ch")) else tadm=tadm+0
rs2.movenext
loop
%>
<tr>
	<td class="campoa" colspan=8><b>Total de horas semanais em atribuições</td>
	<td class="campot" align="center"><%=tadm%></td>
	<td class="campoa" colspan=6></td>
</tr>
<%
end if 'recordcount rs2
rs2.close
%>
<!-- grades outros salarios -->
<%
sql2="SELECT id_rt, chapa, codevento, descricao, ch, chm, inicio, fim " & _
"FROM grades_rt " & _
"WHERE chapa='" & rs("chapa") & "' " & _
"and #" & dtaccess(datad) & "# between inicio and fim " & _
"ORDER BY descricao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<!-- <table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>"> -->
<%
tacad=0
if rs2.recordcount>0 then
%>
<tr><td class=grupo colspan=15>Outras atribuições</td></tr>
<tr>
	<td class=titulo align="center">#</td>
	<td class=titulo align="center" colspan=1>&nbsp;Folha</td>
	<td class=titulo align="center" colspan=6>Descrição</td>
	<td class=titulo align="center">Nr.</td>
	<td class=titulo align="center" colspan=5></td>
	<td class=titulo align="center">Período</td>
</tr>
<%
rs2.movefirst
do while not rs2.eof 
classe="campor"
%>
<tr>
	<td class=<%=classe%> align="center"><%=rs2.absoluteposition%></td>
	<td class=<%=classe%> colspan=1 align="center"><%=rs2("codevento")%></td>
	<td class=<%=classe%> colspan=6><font color=blue><%=rs2("descricao")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=rs2("ch")%></td>
	<td class=<%=classe%> colspan=5>&nbsp;</td>
	<td class=<%=classe%> align="center"><%=rs2("inicio") & " a " & rs2("fim")%></td>
</tr>
<%
if rs2("ch")<>"" then tacad=tacad+cdbl(rs2("ch")) else tacad=tacad+0
rs2.movenext
loop
%>
<tr>
	<td class="campoa" colspan=8><b>Total de horas semanais acadêmicas</td>
	<td class="campot" align="center"><%=tacad%></td>
	<td class="campoa" colspan=6></td>
</tr>
<%
end if 'recordcount rs2
rs2.close
%>

<tr>
	<td class="campot" colspan=8><b>Total da carga horária em horas semanais</td>
	<td class="campoa" align="center"><%=tadm+taulas+tacad%></td>
	<td class="campot" colspan=6></td>
</tr>
</table>
<!----------- inicio frequencia ------------->
<%
sql2="SELECT curso, dia_mes, dia, descr, turma, aulas, faltas, justificativa, atraso, extra, dp, reposicao, observacao " & _
"FROM clc_carga " & _
"WHERE chapa='" & rs("chapa") & "' and dia_mes>=#02/01/05# " & _
"and codcur=41 and turma='1C' " & _
"ORDER BY dia_mes, descr "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<%
if rs2.recordcount>0 then
%>
<br>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr><td class=grupo colspan=11>Frequência e dias de aulas</td></tr>
<tr>
	<td class=titulo align="center">Data</td>
	<td class=titulo align="center">Período</td>
	<td class=titulo align="center">Aulas</td>
	<td class=titulo align="center">Faltas</td>
	<td class=titulo align="center">Justif.</td>
	<td class=titulo align="center">Atraso</td>
	<td class=titulo align="center">Extras</td>
	<td class=titulo align="center">DP</td>
	<td class=titulo align="center">Repos.</td>
	<td class=titulo align="center">Obs.</td>
	<td class=titulo align="center">Marcações eletrônicas</td>
</tr>
<%
rs2.movefirst
do while not rs2.eof 
classe="campor"
if lastcurso<>rs2("curso") then
	response.write "<tr>"
	response.write "<td class="campol" colspan=11><b>" & rs2("curso") & " - Turma: " & rs2("turma") & "</td>"
	response.write "</tr>"
end if

%>
<tr>
	<td class=<%=classe%> align="left"><%=rs2("dia_mes")%>-<%=weekdayname(rs2("dia"),1)%></td>
	<td class=<%=classe%> align="left"><%=rs2("descr")%></td>
	<td class=<%=classe%> align="left"><%=rs2("aulas")%></td>
	<td class=<%=classe%> align="left"><%=rs2("Faltas")%></td>
	<td class=<%=classe%> align="left"><%=rs2("Justificativa")%></td>
	<td class=<%=classe%> align="left"><%=rs2("Atraso")%></td>
	<td class=<%=classe%> align="left"><%=rs2("Extra")%></td>
	<td class=<%=classe%> align="left"><%=rs2("DP")%></td>
	<td class=<%=classe%> align="left"><%=rs2("Reposicao")%></td>
	<td class=<%=classe%> align="left"><%=rs2("Observacao")%></td>
<!-- marcacoes do dia -->
<%
	sqlcr="select chapa, day(data) as dia, data, batida, status from abatfun_m where " & _
	"chapa='" & rs("chapa") & "' and data='" & rs2("dia_mes") & "' order by data, batida"
	sqlcr="select chapa, day(data) as dia, data, batida, status from abatfun where " & _
	"chapa='" & rs("chapa") & "' and data='" & dtaccess(rs2("dia_mes")) & "' order by data, batida"  'sql
	sqlcr="select chapa, day(data) as dia, data, batida, status from abatfun_m where " & _
	"chapa='" & rs("chapa") & "' and data=#" & dtaccess(rs2("dia_mes")) & "# order by data, batida"   'access
	'marcações do chronus
	rs4.Open sqlcr, ,adOpenStatic, adLockReadOnly
	marcacao=0
	for b=1 to 6:marc(b)="":formato(b)="":next
	if rs4.recordcount>0 then
		rs4.movefirst
		do while not rs4.eof
		'dia=rs2("dia")
		batida=formatdatetime((rs4("batida")/60)/24,4)
		'if dia=diaant then 
		marcacao=marcacao+1 'else marcacao=1
		marc(marcacao)="<font color=blue>|</font>" & batida
		if rs4("status")="D" then formato(marcacao)="<font color='red'>" 'else formato(dia,marcacao)="<font color='black'>"
		'diaant=dia
		rs4.movenext
		loop
	else 'recordcount rs2
		for b=1 to 6
			marc(b)=""
		next
	end if 'recordcount rs2
	if marcacao<6 then
		for a=marcacao+1 to 6
			marc(a)=""
		next
	end if
	rs4.close
%>
	<td class=<%=classe%> align="left" style="border-bottom:1px solid #000000">
<%
for a=1 to 6
	response.write formato(a) & marc(a) & "</font>"
next 
response.write "<font color=blue>|</font>"
%>	
	</td>
	
</tr>
<%
lastcurso=rs2("curso")
rs2.movenext
loop
%>
<%
end if 'recordcount rs2
rs2.close
%>
</table>
<%
rs.close
rs3.movenext
response.write "<DIV style=""page-break-after:always""></DIV>"

loop
rs3.close
set rs3=nothing
conexao.close
set conexao=nothing
%>
</BODY>
</HTML>