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
'set conexao2=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
'conexao2.Open Application("consql")
set rs=server.createobject ("ADODB.Recordset")
set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
set rs3.ActiveConnection = conexao
if request("chapa")<>"" then chapap=request("chapa") else chapap=request.form("chapa")
if request("nome")<>"" then chapan=request("nome") else chapan=request.form("nome")

sqla="select f.nome, codsituacao=case when month(datademissao)=month(getdate()) and year(datademissao)=year(getdate()) then 'A' else f.codsituacao end, " & _
"f.chapa, f.dataadmissao, c.nome as funcao, f.codsecao, s.descricao as secao, p.estadocivil, " & _
"p.grauinstrucao, p.rua, p.numero, p.complemento, p.bairro, p.cidade, p.cep, p.telefone1, p.telefone2, p.telefone3, p.corraca, p.deficientefisico, " & _
"p.fax, p.sexo, p.dtnascimento, p.email, p.cpf, p.cartidentidade, p.ufcartident, f.pispasep, f.codpessoa, f.datademissao, " & _
"sit=case when month(datademissao)=month(getdate()) and year(datademissao)=year(getdate()) then 'Ativo' else ss.descricao end, " & _
"i.descricao as titulacao, p.dtemissaoident, pc.titulacaopaga, ic.descricao titulacaopaga, pc.lattes, p.naturalidade, p.estadonatal  " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.pfuncao c, corporerm.dbo.psecao s, corporerm.dbo.pcodsituacao ss, corporerm.dbo.pcodinstrucao i, corporerm.dbo.ppessoa p, corporerm.dbo.pfcompl pc, corporerm.dbo.pcodinstrucao ic " & _
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
if isnull(rs2("lattes")) or rs2("lattes")="" then
	lattes=""
else
	lattes=" (" & _
	"<a class=r href='http://buscatextual.cnpq.br/buscatextual/visualizacv.do?id=" & rs2("lattes") & "' onclick=""NewWindow(this.href,'Lattes','800','600','yes','center');return false"" onfocus=""this.blur()"">" & _
	"Lattes</a>)"
end if
if isnull(rs2("datademissao")) then data2=now() else data2=rs2("datademissao")
ts=datediff("m",rs2("dataadmissao"),data2)
tsa=int(ts/12) : tsm=ts-(tsa*12)
ts=" <b>" & tsa & "</b>A e <b>" & tsm & "</b>M "
%>
<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=titulo>Nome:    </td>
	<td class=titulo></td>
	<td class=titulo>Situação:</td>
	<td class=titulo>Chapa:   </td>
</tr>
<tr>
	<td class=campo><b><%=rs2("nome")%></b></td>
	<td class=campo><%=lattes%></td>
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
	<td class=campo><%=rs2("dataadmissao")%><p style="margin-top:0;margin-bottom:0;font-size:6pt;text-align:center"><%=ts%></p></td>
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
<%
if session("usuariogrupo")<>"COORD.CURSO" and session("usuariogrupo")<>"SECR.GERAL" then
%>
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
<%
end if
%>
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
<%
if session("usuariogrupo")<>"COORD.CURSO" and session("usuariogrupo")<>"SECR.GERAL" then
%>
		<td class=titulo>Nascimento:</td>
<%
end if
%>
		<td class=titulo>E-mail:    </td>
		<td class=titulo>Naturalidade</td>
	</tr>
	<tr>
		<td class=campo><%=rs2("sexo")%></td>
<%
if session("usuariogrupo")<>"COORD.CURSO" and session("usuariogrupo")<>"SECR.GERAL" then
%>
		<td class=campo><%=rs2("dtnascimento")%>&nbsp;&nbsp;&nbsp;&nbsp;(<%=int((now()-rs2("dtnascimento"))/365.25) %>)</td>
<%
end if
%>
		<td class=campo><a href="mailto:<%=rs2("email")%>"><%=rs2("email")%></a></td>
		<td class=campo><%=rs2("naturalidade")%> - <%=rs2("estadonatal")%></td>
	</tr>
	</table>

<%
if session("usuariogrupo")<>"COORD.CURSO" and session("usuariogrupo")<>"SECR.GERAL" then
%>
	<table border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
	<tr>
		<td class=titulo>C.P.F.:</td>
		<td class=titulo>Identidade:</td>
		<td class=titulo>PIS/PASEP:</td>
	</tr>
	<tr>
		<td class=campo><%=rs2("cpf")%></td>
		<td class=campo><!-- <%=rs2("cartidentidade")%> / <%=rs2("ufcartident")%> (<%=rs2("dtemissaoident")%>)-->&nbsp;</td>
		<td class=campo><%=rs2("pispasep")%></td>
	</tr>
	<tr>
<%
sql="SELECT CHAPA, NOME AS MAE FROM corporerm.dbo.PFDEPEND WHERE GRAUPARENTESCO='7' and CHAPA='" & rs2("chapa") & "'"
rs3.open sql, ,adOpenStatic:if rs3.recordcount>0 then mae=rs3("mae") else mae=""
rs3.close

if session("usuariogrupo")="RH" or session("usuariogrupo")="PLANEJAMENTO" then mostraraca=1:colraca=1 else mostraraca=0:colraca=3
sql="select descricao from corporerm.dbo.pcorraca where codcliente=" & rs2("corraca") & " "
rs3.open sql, ,adOpenStatic:if rs3.recordcount>0 then corraca=trim(rs3("descricao"))
rs3.close
%>	
		<td class=titulo colspan=<%=colraca%>>Nome da mãe</td>
<%
if mostraraca=1 then
%>
		<td class=titulo>Cor/Raça</td>
		<td class=titulo>Deficiente?</td>
<%
end if
%>
	</tr>
	<tr>
		<td class=campo colspan=<%=colraca%>><%=mae%></td>
<%
if mostraraca=1 then
%>
		<td class=campo><%=corraca%></td>
		<td class=campo><%if rs2("deficientefisico")="0" or isnull(rs2("deficientefisico")) or rs2("deficientefisico")="" then response.write "<img src='../images/bullet.gif'>" else response.write "<img src='../images/bullet_hl.gif'>"%></td>
<%
end if
%>
	</tr>
    </table>
<%
end if 'session
%>
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
"AND g.deletada=0 and g.chapa1='" & rs2("chapa") & "' and '" & dtaccess(datad) & "' between g.inicio and g.termino " & _
"ORDER BY g.coddoc, g.diasem, g.turno;"
sql2="select g.coddoc, c.curso, g.codtur, g.turno, g.perlet, g.codmat, m.materia, g.diasem, g.chapa1, codsala=max(case when g.codsala is null then t.codsala else g.codsala end), " & _
"g.inicio, g.termino, g.juntar, g.demons, aulas=sum(case when juntar=1 then 0 else case when demons=1 then 0 else 1 end end), jturma=max(g.jturma), " & _
"min(g.horini) horini, max(g.horfim) horfim, t.codcur, t.codper, t.grade " & _
"from g2ch g, corporerm.dbo.umaterias m, g2turmas t, g2cursos c " & _
"where m.codmat collate database_default=g.codmat and t.id_grdturma=g.id_grdturma and t.codcur=c.codcur and t.codper=c.codper " & _
"and g.chapa1='" & rs2("chapa") & "' and '" & dtaccess(datad) & "' between g.inicio and g.termino and g.deletada=0 " & _
"group by g.coddoc, g.codtur, g.turno, g.perlet, g.codmat, m.materia, g.diasem, g.chapa1, g.inicio, g.termino, g.juntar, g.demons, t.codcur, t.codper, t.grade, c.curso " & _
"order by g.coddoc, c.curso, g.diasem, min(g.horini)"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
cortitle="#FFFFFF"
bgtitle="#0000FF"
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr><td class=grupo colspan=13>Grade de atribuição de aulas em 
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
	<td class=titulor align="center" colspan=2 width=60>Horario</td>
	<td class=titulo align="center">Nr.</td>
	<td class=fundor align="center">Alunos</td>
	<td class=titulor align="center">Junta</td>
	<td class=fundor align="center">Consid.</td>
	<td class=titulo align="center">Período</td>
</tr>
<%
rs.movefirst:do while not rs.eof 
aulas=rs("aulas")
'if rs2("turno")="5" then classe="fundor" else classe="campor"
classe="campor"

'if lastcurso<>rs("coddoc") then
if lastcurso<>rs("curso") then
	response.write "<tr>"
	response.write "<td class=""campol"" colspan=13><b>" & rs("curso") & "</td>"
	response.write "</tr>"
end if

sql9="SELECT Count(UMATRICPL.MATALUNO) AS alunos " & _
"FROM corporerm.dbo.USITMAT AS USITMAT_1 INNER JOIN ((corporerm.dbo.UMATRICPL as umatricpl INNER JOIN corporerm.dbo.UMATALUN as umatalun ON (UMATRICPL.CODFILIAL = UMATALUN.CODFILIAL) AND " & _
"(UMATRICPL.CODCOLIGADA = UMATALUN.CODCOLIGADA) AND (UMATRICPL.MATALUNO = UMATALUN.MATALUNO) AND " & _
"(UMATRICPL.PERLETIVO = UMATALUN.PERLETIVO) AND (UMATRICPL.CODCUR = UMATALUN.CODCUR) AND (UMATRICPL.CODPER = UMATALUN.CODPER) " & _
"AND (UMATRICPL.GRADE = UMATALUN.GRADE)) INNER JOIN corporerm.dbo.UMATERIAS as umaterias ON (UMATALUN.CODMAT = UMATERIAS.CODMAT) AND " & _
"(UMATALUN.CODCOLIGADA = UMATERIAS.CODCOLIGADA)) ON USITMAT_1.CODSITMAT = UMATALUN.STATUS " & _
"WHERE (((UMATALUN.STATUS) In ('01','07','08','09','10','15','18','19','20','47','48','46','70','71')) " & _
"AND ((UMATALUN.CODTUR)='" & rs("codtur") & "') " & _
"AND ((UMATRICPL.PERLETIVO)='" & rs("perlet") & "') " & _
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
	<td class=<%=classe%> align="center"><%=rs("turno")%></td>
	
	<td class=<%=classe%> align="center" width=30><%=rs("horini")%></td>
	<td class=<%=classe%> align="center" width=30><%=rs("horfim")%></td>
	
	<td class=<%=classe%> align="center">&nbsp;<%=aulas%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=alunos%></td>

	<td class=<%=classe%> align="center">&nbsp;<%if rs("juntar")=true then response.write "<font face='Wingdings'>ü</font>"%><%=rs("jturma")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%if rs("demons")=false then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center" nowrap><%=rs("inicio") & " a " & rs("termino")%></td>
</tr>
<%
taulas=taulas+aulas
'lastcurso=rs("coddoc")
lastcurso=rs("curso")
rs.movenext
loop
%>
<tr>
	<td class="campoa" colspan=8><b>Total de horas semanais em docência</td>
	<td class="campot" align="center"><%=taulas%></td>
	<td class="campoa" colspan=4></td>
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
"and '" & dtaccess(datad) & "' between mand_ini and mand_fim " & _
"ORDER BY n.nomeacao "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<!-- <table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>"> -->
<%
tadm=0
if rs.recordcount>0 then
%>
<tr><td class=grupo colspan=13>Nomeações e Atividades</td></tr>
<tr>
	<td class=titulo align="center">#</td>
	<td class=titulo align="center" colspan=2>Nomeação</td>
	<td class=titulo align="center" colspan=5>Portaria</td>
	<td class=titulo align="center">Nr.</td>
	<td class=titulo align="center" colspan=2>Curso/cargo/obs.</td>
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
	<td class=<%=classe%> colspan=5><%=rs("portaria")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=rs("ch")%></td>
	<td class=<%=classe%> colspan=2><font color=blue><%=(rs("cargo"))%></td>
	<td class=<%=classe%> align="center"><%if rs("codeve")<>"" then response.write rs("codeve") else response.write "<font color=brown>S/ ônus"%></td>
	<td class=<%=classe%> align="center"><%=rs("mand_ini") & " a " & rs("mand_fim")%></td>
</tr>
<%
if rs("ch")<>"" and rs("codeve")<>"" then tadm=tadm+cdbl(rs("ch")) else tadm=tadm+0
rs.movenext:loop
%>
<tr>
	<td class="campoa" colspan=8><b>Total de horas semanais em atribuições</td>
	<td class="campot" align="center"><%=tadm%></td>
	<td class="campoa" colspan=4></td>
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
"and '" & dtaccess(datad) & "' between inicio and fim " & _
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
	<td class=<%=classe%> colspan=6><font color=blue><%=rs("descricao")%></td>
	<td class=<%=classe%> align="center">&nbsp;<%=rs("ch")%></td>
	<td class=<%=classe%> colspan=3>&nbsp;</td>
	<td class=<%=classe%> align="center"><%=rs("inicio") & " a " & rs("fim")%></td>
</tr>
<%
if rs("ch")<>"" then tacad=tacad+cdbl(rs("ch")) else tacad=tacad+0
rs.movenext:loop
%>
<tr>
	<td class="campoa" colspan=8><b>Total de horas semanais acadêmicas</td>
	<td class="campot" align="center"><%=tacad%></td>
	<td class="campoa" colspan=4></td>
</tr>
<%
end if 'recordcount rs2
rs.close
%>

<tr>
	<td class="campot" colspan=8><b>Total da carga horária em horas semanais</td>
	<td class="campoa" align="center"><%=tadm+taulas+tacad%></td>
	<td class="campot" colspan=4></td>
</tr>
</table>
<%
session("resumida")="N"
if session("resumida")<>"S" then
%>
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

sql2="SELECT f.CODPROF, f.codinstrucao, ft.tipo, f.CURSO, f.INSTITUICAO, f.LOCALINST, f.anoconclusao, f.DATACONCLUSAO, f.ABRANGENCIA, fa.DESCRICAO " & _
"FROM UPROFFORMACAO_ f, UPROF_ABRANGENCIA fa, UPROF_TIPO ft " & _
"WHERE f.ABRANGENCIA=fa.ABRANGENCIA AND f.codinstrucao=ft.codinstrucao AND codprof='" & rs2("chapa") & "' " & _
"ORDER BY f.codinstrucao "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
classe="campor"
if isnull(rs("dataconclusao")) then conclusao=rs("anoconclusao") else conclusao=rs("dataconclusao")
'if indaprov>0 and indaprov<0.75 then classe="campotr" else classe="campor"
%>
<tr>
	<td class=<%=classe%>><%=rs("tipo")%></td>
	<td class=<%=classe%>><%=rs("curso")%></td>
	<td class=<%=classe%>><%=rs("instituicao")%></td>
	<td class=<%=classe%>><%=rs("localinst")%></td>
	<td class=<%=classe%> align="center"><%=conclusao%></td>
	<td class=<%=classe%> ><%=rs("descricao")%></td>
</tr>
<%
rs.movenext:loop
end if
rs.close
%>
</table>

<hr>
<%
if session("usuariogrupo")="RH" or session("usuariomaster")="00440" or session("usuariomaster")="00201" or session("usuariomaster")="00099" then
sqla="SELECT * FROM corporerm.dbo.PANOTAC " & _
"WHERE CODPESSOA=" & rs2("codpessoa") & " and tipo in (19,20) ORDER BY tipo, nroanotacao"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Anotações Pessoais</td></tr>
</table>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr>
	<td class=titulor>Tipo</td>
	<td class=titulor>Data</td>
	<td class=titulor>Texto</td>
	<td class=titulor>Ver</td>
</tr>
<%
rs.movefirst
do while not rs.eof
if rs("tipo")=19 then tipo="Elogio" else tipo="Reclamação"
chamado=trim(right(rs("texto"),10))
chamado2=""
for a=1 to len(chamado)
	letra=mid(chamado,a,1)
	if isnumeric(letra)=true then chamado2=chamado2 & letra
next
chamado=chamado2
%>
<tr>
	<td class="campor" align="left"><%=tipo%></td>
	<td class="campor" align="left"><%=rs("dtanotacao")%></td>
	<td class="campor" align="left"><%=rs("texto")%></td>
	<td class="campor"><a href="http://intranet.unifieo.br/legado/intranet/ouvidoria/visualiza.php?chamado=<%=chamado%>" target="_blank"><%=chamado%></a></td>
</tr>
<%
rs.movenext
loop
%>
</table>
<%
end if
rs.close
end if 'rh ou 00440
%>

<hr>
<%
taprovacao=0 
if taprovacao=1 then ' session("a5")="T" then
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr><td class=grupo colspan=8>Indice de Aprovação de Alunos</td></tr>
<tr>
	<td class=fundo align="center">P.Let.</td>
	<td class=fundo align="center">Turma</td>
	<td class=fundo align="center">Matéria</td>
	<td class=fundo align="center">Aprov.</td>
	<td class=fundo align="center">Rep.Nota</td>
	<td class=fundo align="center">Rep.Freq.</td>
	<td class=fundo align="center">% Aprov</td>
</tr>
<%
sql2="select codcur, curso, perlet, codtur, g.codmat, m.materia, n_aprov, n_repnota, n_repfreq, talunos " & _
"FROM grades_repro g, corporerm.dbo.umaterias m " & _
"WHERE m.codmat collate database_default=g.codmat and chapa1='" & rs2("chapa") & "' and talunos>0  and perlet>convert(varchar(10),year(getdate())-3) " & _
"ORDER BY perlet desc, curso, m.materia "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
classe="campolr"
totalalunos=rs("talunos")
if cint(rs("n_aprov"))>0 then indaprov=rs("n_aprov")/totalalunos else indaprov=0
if indaprov>0.70 and indaprov<0.85 then classe="campotr"
if indaprov>0.50 and indaprov<=0.70 then classe="campoar"
if indaprov>0 and indaprov<=0.5 then classe="camporr"
%>
<tr>
	<td class=<%=classe%>><%=rs("perlet")%></td>
	<td class=<%=classe%>>
	<a class=r href="hstturma.asp?curso=<%=rs("codcur")%>&perlet=<%=rs("perlet")%>&turma=<%=rs("codtur")%>&ncurso=<%=rs("curso")%>" onclick="NewWindow(this.href,'Aprovacao_turma','545','200','yes','center');return false" onfocus="this.blur()">
	<%=rs("codtur")%>
	</a>
	</td>
	<td class=<%=classe%>><%=rs("materia")%></td>
	<td class=<%=classe%> align="center"><%=rs("n_aprov")%></td>
	<td class=<%=classe%> align="center"><%=rs("n_repnota")%></td>
	<td class=<%=classe%> align="center"><%=rs("n_repfreq")%></td>
	<td class=<%=classe%> align="center"><%=formatpercent(indaprov,2)%></td>
</tr>
<%
rs.movenext
loop
end if
rs.close
%>
</table>
<% end if 'grantdocens para ver aprovacao %>
<br>
<%
if session("usuariomaster")="88888" or session("usuariomaster")="02379" or session("usuariogrupo")="99" or session("usuariomaster")="02653" OR session("usuariomaster")="00259" then
chapa=rs2("chapa")
sqlmeses="SELECT DISTINCT TOP 6 CHAPA, ANOCOMP, MESCOMP, convert(datetime, str(anocomp)+'/'+str(mescomp)+'/'+str(1)) as dc " & _
"FROM (select chapa, anocomp, mescomp from corporerm.dbo.PFPERFF union all select chapa, anocomp, mescomp from corporerm.dbo.PFPERFFcompl) f " & _
"WHERE CHAPA='" & chapa & "' ORDER BY ANOCOMP DESC, MESCOMP DESC"
rs3.CursorLocation = adUseClient
rs3.Open sqlmeses, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
	pos=0:rs3.movelast
	do while not rs3.bof
		pos=pos+1
		stringmeses=stringmeses & "sum(case dc when '" & dtaccess(rs3("dc")) & "' then ff.valor else 0 end) '" & month(rs3("dc")) & "/" & year(rs3("dc")) & "'"
		if pos<rs3.recordcount then stringmeses=stringmeses & ", "
	rs3.moveprevious:loop
end if
rs3.close
'response.write stringmeses & "<br><br>"
if stringmeses<>"" then juncao="," else juncao=""

sqlficha="SELECT ff.CODEVENTO, e.DESCRICAO " &  juncao & stringmeses & _
"FROM ((select * from corporerm.dbo.PFFINANC union all select * from corporerm.dbo.PFFINANCCOMPL) AS ff INNER JOIN corporerm.dbo.PEVENTO AS e ON ff.CODEVENTO=e.CODIGO) INNER JOIN " & _
"(SELECT DISTINCT TOP 6 CHAPA, ANOCOMP, MESCOMP, convert(datetime, str(anocomp)+'/'+str(mescomp)+'/'+str(1)) as dc FROM (select chapa, anocomp, mescomp from corporerm.dbo.PFPERFF union all select chapa, anocomp, mescomp from corporerm.dbo.PFPERFFcompl) f " & _
"WHERE CHAPA='" & chapa & "' ORDER BY ANOCOMP DESC, MESCOMP DESC) sel ON (ff.MESCOMP=sel.MESCOMP) AND (ff.ANOCOMP=sel.ANOCOMP) AND (ff.CHAPA=sel.CHAPA) " & _
"WHERE E.PROVDESCBASE='P' and ff.valor>0 " & _
"GROUP BY ff.NROPERIODO, ff.CHAPA, ff.CODEVENTO, e.DESCRICAO " & _
"ORDER BY ff.NROPERIODO, ff.CODEVENTO "
'response.write sqlficha
rs3.Open sqlficha, ,adOpenStatic, adLockReadOnly

'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs3.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs3.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'do while not rs3.eof 
'response.write "<tr>"
'for a= 0 to rs3.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs3.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs3.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************
%>
<p class=realce>Ficha Financeira</p>
<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse' width=100%>
<tr>
	<td class=titulo>Cod.</td>
	<td class=titulo align="center">Descrição</td>
	<%for a=2 to rs3.fields.count-1%>
		<td class=titulo align="center"><%=rs3.fields(a).name%></td>
	<%next%>
</tr>
<%
if rs3.recordcount=0 then
%>
<tr>
	<td class=campo colspan=8><b>Sem rendimentos nos ultimos meses</td>
</tr>
<%
else
	rs3.movefirst
	do while not rs3.eof
%>
<tr>
	<td class=campo><%=rs3("codevento")%></td>
	<td class=campo><%=rs3("descricao")%></td>
	<%for a=2 to rs3.fields.count-1%>
		<td class=campo align="right"><%=formatnumber(rs3.fields(a),2)%></td>
	<%next%>
</tr>
<%
	rs3.movenext
	loop
end if	
	rs3.close


sqlficha="SELECT 'Sub-Total', 'Sub-Total', " & stringmeses & _
"FROM ((select * from corporerm.dbo.PFFINANC union all select * from corporerm.dbo.PFFINANCCOMPL) AS ff INNER JOIN corporerm.dbo.PEVENTO AS e ON ff.CODEVENTO=e.CODIGO) INNER JOIN " & _
"(SELECT DISTINCT TOP 6 CHAPA, ANOCOMP, MESCOMP, convert(datetime, str(anocomp)+'/'+str(mescomp)+'/'+str(1)) as dc FROM (select chapa, anocomp, mescomp from corporerm.dbo.PFPERFF union all select chapa, anocomp, mescomp from corporerm.dbo.PFPERFFcompl) f " & _
"WHERE CHAPA='" & chapa & "' ORDER BY ANOCOMP DESC, MESCOMP DESC) sel ON (ff.MESCOMP=sel.MESCOMP) AND (ff.ANOCOMP=sel.ANOCOMP) AND (ff.CHAPA=sel.CHAPA) " & _
"WHERE E.PROVDESCBASE='P' and ff.valor>0 " 
'response.write sqlficha
if stringmeses<>"" then
	rs3.Open sqlficha, ,adOpenStatic, adLockReadOnly
	%>
	<tr>
		<td class=campo colspan=2>TOTAL</td>
		<%for a=2 to rs3.fields.count-1
		if rs3.recordcount=0 then total=0 else total=rs3.fields(a)
		if isnull(total) then total=0
		%>
			<td class=fundo align="right"><b><%=formatnumber(total,2)%></td>
		<%next%>
	</tr>
	<%
	rs3.close
end if
%>
</table>
<% 

end if 'request.resumida

end if 'session(usuariomaster)
%>

</form>
<%
rs2.close
set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
'conexao2.close
set conexao=nothing
'set conexao2=nothing
%>
</body>
</html>