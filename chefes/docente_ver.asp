<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
'revisado com divisao de codigos de cursos
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a5")="N" or session("a5")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"

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
set rs2.ActiveConnection = conexao2
set rs3=server.createobject ("ADODB.Recordset")
set rs3.ActiveConnection = conexao2
if request("chapa")<>"" then chapap=request("chapa") else chapap=request.form("chapa")
if request("nome")<>"" then chapan=request("nome") else chapan=request.form("nome")

sqla="select f.nome, codsituacao=case when month(datademissao)=month(getdate()) and year(datademissao)=year(getdate()) then 'A' else f.codsituacao end, " & _
"f.chapa, f.dataadmissao, c.nome as funcao, f.codsecao, s.descricao as secao, p.estadocivil, " & _
"p.grauinstrucao, p.rua, p.numero, p.complemento, p.bairro, p.cidade, p.cep, p.telefone1, p.telefone2, p.telefone3, p.corraca, p.deficientefisico, " & _
"p.fax, p.sexo, p.dtnascimento, p.email, p.cpf, p.cartidentidade, p.ufcartident, f.pispasep, f.codpessoa, f.datademissao, " & _
"sit=case when month(datademissao)=month(getdate()) and year(datademissao)=year(getdate()) then 'Ativo' else ss.descricao end, " & _
"i.descricao as titulacao, p.dtemissaoident, pc.titulacaopaga, ic.descricao titulacaopaga  " & _
"from pfunc f, pfuncao c, psecao s, pcodsituacao ss, pcodinstrucao i, ppessoa p, pfcompl pc, pcodinstrucao ic " & _
"where f.codfuncao=c.codigo and f.codsecao=s.codigo and p.codigo=f.codpessoa and pc.chapa=f.chapa and pc.titulacaopaga=ic.codcliente " & _
"and f.codsituacao=ss.codcliente and p.grauinstrucao=i.codcliente "
sqlb="and f.CHAPA='" & request("chapa") & "' "
sqlc="ORDER BY f.CHAPA"
sql1=sqla & sqlb & sqlc
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
if request.form("datad")<>"" then datad=formatdatetime(request.form("datad"),2) else datad=formatdatetime(now(),2)
if request.form("datad")="" and rs2("codsituacao")="D" then datad=rs2("datademissao")
%>
<form method="POST" name="form" action="docente_ver.asp">
<p style="margin-top: 0; margin-bottom: 0" class=titulo>CADASTRO DE DOCENTES</p>
<input type="hidden" name="chapa" value="<%=rs2("chapa")%>">
<input type="hidden" name="nome" value="<%=rs2("nome")%>">
<%
session("chapa")=rs2("chapa")
session("chapanome")=rs2("nome")
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
	<td class=campo><b><%=rs2("nome")%></b></td>
	<td class=campo><%=rs2("sit")%>&nbsp;<%if rs2("codsituacao")="D" then response.write rs2("datademissao")%></td>
	<td class=campo><%=rs2("chapa")%>      </td>
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
	<td class=campo><%=rs2("dataadmissao")%></td>
	<td class=campo><%=rs2("funcao")%>      </td>
	<td class=campo><%=rs2("codsecao")%>&nbsp;<%=rs2("secao")%></td>
	<td class=campo>MEC/Paga: <b><%=rs2("titulacaopaga")%></b><br>Real: <b><%=rs2("titulacao")%></td>
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
		<td class=titulo>Endereço:</td>
		<td class=titulo>Complemento:</td>
	</tr>
	<tr>
		<td class=campo><%=rs2("rua") & ", " & rs2("numero")%></td>
		<td class=campo><%=rs2("complemento")%></td>
	</tr>
	</table>

	<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
	<tr>
		<td class=titulo>Bairro:</td>
		<td class=titulo>Cidade:</td>
		<td class=titulo>CEP:</td>
	</tr>
	<tr>
		<td class=campo><%=rs2("bairro")%></td>
		<td class=campo><%=rs2("cidade")%></td>
		<td class=campo><%=rs2("cep")%>   </td>
	</tr>
	</table>

	<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
	<tr>
		<td class=titulo>Telefone 1:</td>
		<td class=titulo>Telefone 2:</td>
		<td class=titulo>Telefone 3:</td>
		<td class=titulo>Fax:</td>
	</tr>
	<tr>
		<td class=campo><%=rs2("telefone1")%></td>
		<td class=campo><%=rs2("telefone2")%></td>
		<td class=campo><%=rs2("telefone3")%></td>
		<td class=campo><%=rs2("fax")%>      </td>
	</tr>
	</table>

	<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
	<tr>
		<td class=titulo>Sexo:      </td>
		<td class=titulo>Nascimento:</td>
		<td class=titulo>E-mail:    </td>
	</tr>
	<tr>
		<td class=campo><%=rs2("sexo")%></td>
		<td class=campo><%=rs2("dtnascimento")%>&nbsp;&nbsp;&nbsp;&nbsp;(<%=int((now()-rs2("dtnascimento"))/365.25) %>)</td>
		<td class=campo><a href="mailto:<%=rs2("email")%>"><%=rs2("email")%></a></td>
	</tr>
	</table>

    </td>
	<td width="<%=tbfoto%>" valign="top">
		<img border="0" src="func_foto.asp?chapa=<%=rs2("chapa")%>"  width="<%=tbfoto%>">
	</td>
</tr>
</table>

<!-- grades horaria de aulas -->
<%
'rs.movenext:loop
sql2="SELECT g.*, c.CURSO AS curso2, p.codcur " & _
"FROM grades_gc gc, g2cursoeve c, g2ch g, (select coddoc, enfase, perlet, gc, codcur, curso, perlet2, perletsg from grades_per group by coddoc, enfase, perlet, gc, codcur, curso, perlet2, perletsg) as p " & _
"WHERE c.coddoc=g.coddoc AND (p.perletsg=g.perletsg AND p.perlet2=g.perlet2 AND p.perlet=g.perlet AND c.coddoc=p.coddoc) " & _
"AND (gc.serie=g.serie AND gc.GC=p.gc AND gc.perlet=p.perlet AND gc.coddoc= p.coddoc) " & _
"AND g.deletada=0 and g.chapa1='" & rs2("chapa") & "' and #" & dtaccess(datad) & "# between g.inicio and g.termino " & _
"ORDER BY g.coddoc, g.diasem, g.turno;"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
cortitle="#FFFFFF"
bgtitle="#0000FF"
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr><td class=grupo colspan=19>Grade de atribuição de aulas em 
	<input type="text" name="datad" size="12" maxlength="12" value="<%=datad%>" class=form_apt onChange="javascript:submit()">
	</td></tr>
<%
taulas=0
if rs.recordcount>0 then
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
	<td class=titulor align="center">Extra</td>
	<td class=fundor align="center">Consid.</td>
	<td class=titulo align="center">Período</td>
</tr>
<%
rs.movefirst:do while not rs.eof 
aulas=rs("ta")
'if rs2("turno")="5" then classe="fundor" else classe="campor"
classe="campor"
if rs("a1")=1 then classe1="fundor" else classe1="campor"
if rs("a2")=1 then classe2="fundor" else classe2="campor"
if rs("a3")=1 then classe3="fundor" else classe3="campor"
if rs("a4")=1 then classe4="fundor" else classe4="campor"
if rs("a5")=1 then classe5="fundor" else classe5="campor"
if rs("a6")=1 then classe6="fundor" else classe6="campor"

if lastcurso<>rs("coddoc") then
	response.write "<tr>"
	response.write "<td class="campol" colspan=19><b>" & rs("curso2") & "</td>"
	response.write "</tr>"
end if
if rs("turno")="1" then turno="Mat"
if rs("turno")="2" then turno="Vesp"
if rs("turno")="3" then turno="Not"
if rs("turno")="5" then turno="Vesp-EF"
if rs("turno")="61" then turno="Int.M"
if rs("turno")="62" then turno="Int.V"
if rs("turno")="6" then turno="Int."

sql9="SELECT Count(UMATRICPL.MATALUNO) AS alunos " & _
"FROM USITMAT AS USITMAT_1 INNER JOIN ((UMATRICPL INNER JOIN UMATALUN ON (UMATRICPL.CODFILIAL = UMATALUN.CODFILIAL) AND " & _
"(UMATRICPL.CODCOLIGADA = UMATALUN.CODCOLIGADA) AND (UMATRICPL.MATALUNO = UMATALUN.MATALUNO) AND " & _
"(UMATRICPL.PERLETIVO = UMATALUN.PERLETIVO) AND (UMATRICPL.CODCUR = UMATALUN.CODCUR) AND (UMATRICPL.CODPER = UMATALUN.CODPER) " & _
"AND (UMATRICPL.GRADE = UMATALUN.GRADE)) INNER JOIN UMATERIAS ON (UMATALUN.CODMAT = UMATERIAS.CODMAT) AND " & _
"(UMATALUN.CODCOLIGADA = UMATERIAS.CODCOLIGADA)) ON USITMAT_1.CODSITMAT = UMATALUN.STATUS " & _
"WHERE (((UMATALUN.STATUS) In ('01','07','08','09','10','15','18','19','20','47','48','46','70','71')) " & _
"AND ((UMATALUN.CODTUR)='" & rs("codtur") & "') " & _
"AND ((UMATRICPL.PERLETIVO)='" & rs("perletsg") & "') " & _
"AND ((UMATRICPL.CODCUR)=" & rs("codcur")& ") " & _
"AND ((UMATALUN.CODMAT)='" & rs("codmat") & "')) "
rs3.Open sql9, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then alunos=rs3("alunos") else alunos="-"
rs3.close
%>
<tr>
	<td class=<%=classe%> align="center"><%=rs.absoluteposition%></td>
	<td class=<%=classe%> align="center"><%=rs("perlet")%></td>
	<td class=<%=classe%> ><font color=blue><%=rs("materia")%></td>
	<td class=<%=classe%> align="center" nowrap><%=rs("codtur")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=weekdayname(rs("diasem"),1)%></td>
	<td class=<%=classe%> align="center"><%=turno%></td>
	
	<td class=<%=classe1%> align="center" width=10><%=rs("a1")%></td>
	<td class=<%=classe2%> align="center" width=10><%=rs("a2")%></td>
	<td class=<%=classe3%> align="center" width=10><%=rs("a3")%></td>
	<td class=<%=classe4%> align="center" width=10><%=rs("a4")%></td>
	<td class=<%=classe5%> align="center" width=10><%=rs("a5")%></td>
	<td class=<%=classe6%> align="center" width=10><%=rs("a6")%></td>
	
	<td class=<%=classe%> align="center">&nbsp;<%=aulas%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=alunos%></td>

	<td class=<%=classe%> align="center">&nbsp;<%if rs("juntar")=true then response.write "<font face='Wingdings'>ü</font>"%><%=rs("jturma")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%if rs("dividir")=true then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center">&nbsp;<%if rs("extra")=true then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center">&nbsp;<%if rs("demons")=false then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center" nowrap><%=rs("inicio") & " a " & rs("termino")%></td>
</tr>
<%
taulas=taulas+aulas
lastcurso=rs("coddoc")
rs.movenext
loop
%>
<tr>
	<td class="campoa" colspan=12><b>Total de horas semanais em docência</td>
	<td class="campot" align="center"><%=taulas%></td>
	<td class="campoa" colspan=6></td>
</tr>
<%
end if 'recordcount rs2
rs.close
%>
<!-- </table> -->

<!-- grades atribuicoes/nomeacoes -->
<%
sql2="SELECT i.id_nomeacao, n.NOMEACAO, i.id_indicado, i.CHAPA, i.NOME, i.PORTARIA, " & _
"i.codeve, i.CARGO, i.MAND_INI, i.MAND_FIM, i.CH, i.obs, i.contrato " & _
"FROM n_indicacoes as i, n_nomeacoes as n " & _
"WHERE i.id_nomeacao = n.id_nomeacao and i.chapa='" & rs2("chapa") & "' " & _
"and #" & dtaccess(datad) & "# between mand_ini and mand_fim " & _
"ORDER BY n.nomeacao "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<!-- <table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>"> -->
<%
tadm=0
if rs.recordcount>0 then
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
rs.movefirst:do while not rs.eof 
classe="campor"
%>
<tr>
	<td class=<%=classe%> align="center"><%=rs.absoluteposition%></td>
	<td class=<%=classe%> colspan=2><%=(rs("nomeacao"))%></td>
	<td class=<%=classe%> colspan=9><font color=blue><%=(rs("cargo"))%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=rs("ch")%></td>
	<td class=<%=classe%> colspan=4><%=rs("portaria")%></td>
	<td class=<%=classe%> align="center"><%if rs("codeve")<>"" then response.write rs("codeve") else response.write "<font color=brown>S/ ônus"%></td>
	<td class=<%=classe%> align="center"><%=rs("mand_ini") & " a " & rs("mand_fim")%></td>
</tr>
<%
if rs("ch")<>"" and rs("codeve")<>"" then tadm=tadm+cdbl(rs("ch")) else tadm=tadm+0
rs.movenext:loop
%>
<tr>
	<td class="campoa" colspan=12><b>Total de horas semanais em atribuições</td>
	<td class="campot" align="center"><%=tadm%></td>
	<td class="campoa" colspan=6></td>
</tr>
<%
end if 'recordcount rs2
rs.close
%>
<!-- grades outros salarios -->
<%
sql2="SELECT id_rt, chapa, codevento, descricao, ch, chm, inicio, fim " & _
"FROM grades_rt " & _
"WHERE chapa='" & rs2("chapa") & "' " & _
"and #" & dtaccess(datad) & "# between inicio and fim " & _
"ORDER BY descricao "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<!-- <table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>"> -->
<%
tacad=0
if rs.recordcount>0 then
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
rs.movefirst:do while not rs.eof 
classe="campor"
%>
<tr>
	<td class=<%=classe%> align="center"><%=rs.absoluteposition%></td>
	<td class=<%=classe%> colspan=1 align="center"><%=rs("codevento")%></td>
	<td class=<%=classe%> colspan=10><font color=blue><%=rs("descricao")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=rs("ch")%></td>
	<td class=<%=classe%> colspan=5>&nbsp;</td>
	<td class=<%=classe%> align="center"><%=rs("inicio") & " a " & rs("fim")%></td>
</tr>
<%
if rs("ch")<>"" then tacad=tacad+cdbl(rs("ch")) else tacad=tacad+0
rs.movenext:loop
%>
<tr>
	<td class="campoa" colspan=12><b>Total de horas semanais acadêmicas</td>
	<td class="campot" align="center"><%=tacad%></td>
	<td class="campoa" colspan=6></td>
</tr>
<%
end if 'recordcount rs2
rs.close
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
"WHERE f.ABRANGENCIA=fa.ABRANGENCIA AND f.TIPO=ft.TIPO AND codprof='" & rs2("chapa") & "' " & _
"ORDER BY f.CODPROF, ft.ID "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
classe="campor"
'if indaprov>0 and indaprov<0.75 then classe="campotr" else classe="campor"
%>
<tr>
	<td class=<%=classe%>><%=rs("tipo")%></td>
	<td class=<%=classe%>><%=rs("curso")%></td>
	<td class=<%=classe%>><%=rs("instituicao")%></td>
	<td class=<%=classe%>><%=rs("localinst")%></td>
	<td class=<%=classe%> align="center"><%=rs("dataconclusao")%></td>
	<td class=<%=classe%> ><%=rs("descricao")%></td>
</tr>
<%
rs.movenext:loop
end if
rs.close
%>
</table>

<hr>

</form>
<%
rs2.close
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