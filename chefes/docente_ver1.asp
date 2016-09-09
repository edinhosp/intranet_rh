<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
'revisado com divisao de codigos de cursos
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a98")="N" or session("a98")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"

if request("nome")<>"" then chapan=request("nome") else chapan=request.form("nome")
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Docentes - <%=chapan%></title>
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

dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
set conexao2=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
conexao2.Open Application("consql")
set rs=server.createobject ("ADODB.Recordset")
set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
set rs3.ActiveConnection = conexao2
if request("chapa")<>"" then chapap=request("chapa") else chapap=request.form("chapa")
if request("nome")<>"" then chapan=request("nome") else chapan=request.form("nome")

sqla="select f.nome, iif(month(datademissao)=month(now) and year(datademissao)=year(now),'A',f.codsituacao) as codsituacao, f.chapa, f.dataadmissao, c.nome as funcao, f.codsecao, s.descricao as secao, f.estadocivil, " & _
"f.grauinstrucao, f.rua, f.numero, f.complemento, f.bairro, f.cidade, f.cep, f.telefone1, f.telefone2, f.telefone3, f.corraca, f.deficientefisico, " & _
"f.fax, f.sexo, f.dtnascimento, f.email, f.cpf, f.cartidentidade, f.ufcartident, f.pispasep, f.codpessoa, f.datademissao, " & _
"iif(month(datademissao)=month(now) and year(datademissao)=year(now),'Ativo',ss.descricao) as sit, i.descricao as titulacao, f.dtemissaoident, f.mae " & _
"from dc_professor_t f, pfuncao c, psecao s, pcodsituacao ss, pcodinstrucao i " & _
"where f.codfuncao = c.codigo and f.codsecao = s.codigo " & _
"and f.codsituacao = ss.codcliente and f.grauinstrucao=i.codcliente "
sqlb="and f.CHAPA='" & request("chapa") & "' "
sqlc="ORDER BY f.CHAPA"
sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if request.form("datad")<>"" then datad=formatdatetime(request.form("datad"),2) else datad=formatdatetime(now(),2)
if request.form("datad")="" and rs("codsituacao")="D" then datad=rs("datademissao")
%>
<form method="POST" name="form" action="docente_ver.asp">
<p style="margin-top: 0; margin-bottom: 0" class=titulo>CADASTRO DE DOCENTES</p>
<input type="hidden" name="chapa" value="<%=rs("chapa")%>">
<input type="hidden" name="nome" value="<%=rs("nome")%>">
<%
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
tabela=615
tbfoto=150
%>
<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela%>">
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

<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Dados Acadêmicos</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela%>">
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

<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela%>">
<tr><td class=grupo>Dados Pessoais</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td valign="top" class=campo>

	<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
	<tr>
		<td class=titulo>Telefone 1:</td>
		<td class=titulo>Telefone 2:</td>
		<td class=titulo>Telefone 3:</td>
		<td class=titulo>Fax:</td>
	</tr>
	<tr>
		<td class=campo><%=rs("telefone1")%></td>
		<td class=campo><%=rs("telefone2")%></td>
		<td class=campo><%=rs("telefone3")%></td>
		<td class=campo><%=rs("fax")%>      </td>
	</tr>
	</table>

	<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
	<tr>
		<td class=titulo>Sexo:      </td>
		<td class=titulo>Nascimento:</td>
		<td class=titulo>E-mail:    </td>
	</tr>
	<tr>
		<td class=campo><%=rs("sexo")%></td>
		<td class=campo><%=rs("dtnascimento")%>&nbsp;&nbsp;&nbsp;&nbsp;(<%=int((now()-rs("dtnascimento"))/365.25) %>)</td>
		<td class=campo><a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></td>
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

sql2="select * from g2ch where deletada=0 and chapa1='" & rs("chapa") & "' " & _
"and #" & dtaccess(datad) & "# between inicio and termino order by curso, diasem, turno "
sql2="SELECT g.*, gc.coddoc, gc.CURSO as curso2 " & _
"FROM g2ch AS g INNER JOIN (g2cursoeve AS gc INNER JOIN g2cursoeve_grupo AS gcg ON gc.coddoc=gcg.doc) ON (g.codcur=gcg.codcur) AND (g.enfase=gcg.enfase) " & _
"WHERE g.deletada=0 and g.chapa1='" & rs("chapa") & "' " & _
"and #" & dtaccess(datad) & "# between g.inicio and g.termino " & _
"ORDER BY gc.coddoc, g.diasem, g.turno; "
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
	<td class=titulor align="center" colspan=6 width=60>Horario</td>
	<td class=titulo align="center">Nr.</td>
	<td class=fundor align="center">Alunos</td>
	<td class=titulor align="center">Junta</td>
	<td class=titulor align="center">Divide</td>
</tr>
<%
rs2.movefirst
do while not rs2.eof 
aulas=rs2("ta")
'if rs2("turno")="5" then classe="fundor" else classe="campor"
classe="campor"
if rs2("a1")=1 then classe1="fundor" else classe1="campor"
if rs2("a2")=1 then classe2="fundor" else classe2="campor"
if rs2("a3")=1 then classe3="fundor" else classe3="campor"
if rs2("a4")=1 then classe4="fundor" else classe4="campor"
if rs2("a5")=1 then classe5="fundor" else classe5="campor"
if rs2("a6")=1 then classe6="fundor" else classe6="campor"

if lastcurso<>rs2("coddoc") then
	response.write "<tr>"
	response.write "<td class="campol" colspan=19><b>" & rs2("curso2") & "</td>"
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
"WHERE (((UMATALUN.STATUS) In ('01','07','08','09','10','15','18','19','20','47','48','46','70','71')) " & _
"AND ((UMATALUN.CODTUR)='" & rs2("codtur") & "') " & _
"AND ((UMATRICPL.PERLETIVO)='" & rs2("perletsg") & "') " & _
"AND ((UMATRICPL.CODCUR)=" & rs2("codcur")& ") " & _
"AND ((UMATALUN.CODMAT)='" & rs2("codmat") & "')) "
rs3.Open sql9, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then alunos=rs3("alunos") else alunos="-"
rs3.close
%>
<tr>
	<td class=<%=classe%> align="center"><%=rs2.absoluteposition%></td>
	<td class=<%=classe%> align="center"><%=rs2("perlet")%></td>
	<td class=<%=classe%> ><font color=blue><%=rs2("materia")%></td>
	<td class=<%=classe%> align="center" nowrap><%=rs2("codtur")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=weekdayname(rs2("diasem"),1)%></td>
	<td class=<%=classe%> align="center"><%=turno%></td>
	
	<td class=<%=classe1%> align="center" width=10><%=rs2("a1")%></td>
	<td class=<%=classe2%> align="center" width=10><%=rs2("a2")%></td>
	<td class=<%=classe3%> align="center" width=10><%=rs2("a3")%></td>
	<td class=<%=classe4%> align="center" width=10><%=rs2("a4")%></td>
	<td class=<%=classe5%> align="center" width=10><%=rs2("a5")%></td>
	<td class=<%=classe6%> align="center" width=10><%=rs2("a6")%></td>
	
	<td class=<%=classe%> align="center">&nbsp;<%=aulas%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=alunos%></td>

	<td class=<%=classe%> align="center">&nbsp;<%if rs2("juntar")=true then response.write "<font face='Wingdings'>ü</font>"%><%=rs2("jturma")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%if rs2("dividir")=true then response.write "<font face='Wingdings'>ü</font>" %></td>
</tr>
<%
taulas=taulas+aulas
lastcurso=rs2("coddoc")
rs2.movenext
loop
%>
<tr>
	<td class="campoa" colspan=12><b>Total de horas semanais em docência</td>
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
<tr><td class=grupo colspan=19>Nomeações e Atividades</td></tr>
<tr>
	<td class=titulo align="center">#</td>
	<td class=titulo align="center" colspan=2>Nomeação</td>
	<td class=titulo align="center" colspan=9>Curso/cargo/obs.</td>
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
	<td class=<%=classe%> colspan=9><font color=blue><%=(rs2("cargo"))%></td>
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
	<td class="campoa" colspan=12><b>Total de horas semanais em atribuições</td>
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
<tr><td class=grupo colspan=19>Outras atribuições</td></tr>
<tr>
	<td class=titulo align="center">#</td>
	<td class=titulo align="center" colspan=1>&nbsp;Folha</td>
	<td class=titulo align="center" colspan=10>Descrição</td>
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
	<td class=<%=classe%> colspan=10><font color=blue><%=rs2("descricao")%></td>
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
	<td class="campoa" colspan=12><b>Total de horas semanais acadêmicas</td>
	<td class="campot" align="center"><%=tacad%></td>
	<td class="campoa" colspan=6></td>
</tr>
<%
end if 'recordcount rs2
rs2.close
%>

<tr>
	<td class="campot" colspan=12><b>Total da carga horária em horas semanais</td>
	<td class="campoa" align="center"><%=tadm+taulas+tacad%></td>
	<td class="campot" colspan=6></td>
</tr>
</table>

<hr>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr><td class=grupo colspan=6>Formação Acadêmica</td></tr>
<tr>
	<td class=fundo align="center">Tipo</td>
	<td class=fundo align="center">Curso</td>
	<td class=fundo align="center">Instituição</td>
	<td class=fundo align="center">Local</td>
	<td class=fundo align="center">Ano</td>
	<td class=fundo align="center">Abrangência</td>
</tr>
<%
sql2="SELECT f.CODPROF, ft.ID, f.TIPO, f.CURSO, f.INSTITUICAO, f.LOCALINST, f.DATACONCLUSAO, f.ABRANGENCIA, fa.DESCRICAO " & _
"FROM UPROFFORMACAO_ f, UPROF_ABRANGENCIA fa, UPROF_TIPO ft " & _
"WHERE f.ABRANGENCIA=fa.ABRANGENCIA AND f.TIPO=ft.TIPO AND codprof='" & rs("chapa") & "' " & _
"ORDER BY f.CODPROF, ft.ID "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
classe="campor"
'if indaprov>0 and indaprov<0.75 then classe="campotr" else classe="campor"
%>
<tr>
	<td class=<%=classe%>><%=rs2("tipo")%></td>
	<td class=<%=classe%>><%=rs2("curso")%></td>
	<td class=<%=classe%>><%=rs2("instituicao")%></td>
	<td class=<%=classe%>><%=rs2("localinst")%></td>
	<td class=<%=classe%> align="center"><%=rs2("dataconclusao")%></td>
	<td class=<%=classe%> ><%=rs2("descricao")%></td>
</tr>
<%
rs2.movenext
loop
end if
rs2.close
%>
</table>

<hr>

</form>
<%
rs.close
set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
conexao2.close
set conexao=nothing
set conexao2=nothing
%>
</body>
</html>