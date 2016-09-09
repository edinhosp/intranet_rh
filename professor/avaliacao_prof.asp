<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a89")="N" or session("a89")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Emissão de Avaliação de Professor Recem-Contratado</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"><!--
function nome1() {	form.chapa.value=form.nome.value; }
function chapa1() {	form.nome.value=form.chapa.value; }
--></script>
</head>
<body style="margin-left:20px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sessao=session.sessionid
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

espacamento=5
if request.form="" then
sql="select p.chapa, p.nome, p.dataadmissao from corporerm.dbo.pfunc p where p.chapa<'10000' and p.codtipo='N' and codsituacao<>'D' and codsindicato='03' " & _
"order by p.dataadmissao desc, p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="form" action="avaliacao_prof.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleção de Professor para emissão da Avaliação</td>
</tr>
<tr>
	<td class=campo>Professor</td>
	<td class=campo>
		<select name="chapa" class=a size=20 multiple>
		<option value="0"> Selecione o professor</option>
<%
rs.movefirst
do while not rs.eof
%>
		<option value="<%=rs("chapa")%>"> <%=rs("chapa") & " - " & rs("nome") & " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(" & rs("dataadmissao") & ")"%></option>
<%
rs.movenext
loop
rs.close
%>
		</select>
	</td>
</tr>
<tr>
	<td class=campo colspan=3>&nbsp;
	<input type="radio" name="tipo" value="T" checked>Tudo
	<input type="radio" name="tipo" value="C">Para Coordenador
	<input type="radio" name="tipo" value="A">Para Auto-Avaliação
	</td>
</tr>
<tr>
	<td class=campo colspan=3>&nbsp;
		<input type="submit" value="Visualizar" class=button name="B1">
	</td>
</tr>
</table>
</form>

<%
else
'response.write request.form("chapa").count
chapas=request.form("chapa").count
tipo=request.form("tipo")

sql="delete from ttavalprof where sessao='" & session("usuariomaster") & "'"
conexao.execute sql
if tipo="T" or tipo="C" then
%>
<table border="0" cellpadding="1" cellspacing="0" width="930" height=470 style="border-collapse: collapse">
<%
'****
dim perg(14)
perg(1)="Possui boa apresentação pessoal"
perg(2)="Teve fácil adaptação e respeito às normas da Instituição"
perg(3)="Costuma chegar e sair da sala de aula no horário correto"
perg(4)="Apresenta assiduidade, comparece para ministrar aula, raramente falta"
perg(5)="Cumpre o plano de ensino e segue as bibliografias indicadas"
perg(6)="Apresenta conteúdos atualizados e utiliza sempre que necessário recursos áudio-visuais em suas aulas"
perg(7)="Apresenta didática e conhecimentos necessários para transmissão dos conteúdos aos alunos"
perg(8)="Cumpre prazos para entrega e publicação das notas de avaliações e/ou trabalhos"
perg(9)="Relaciona-se de forma positiva com os alunos (disciplina, respeito e confiança)"
perg(10)="Os alunos estão satisfeitos com sua forma de lecionar. Não existe registros de queixas na coordenação e ouvidoria"
perg(11)="Tem facilidade de relacionamento com os colegas e coordenador"
perg(12)="Está disposto a colaborar com os demais, tem iniciativa e entusiasmo para trabalhos coletivos"
perg(13)="Participa nas reuniões de trabalho, demonstrando interesse e envolvimento"
perg(14)="Indicaria este professor para outro curso ou disciplina"

for a=1 to chapas
'****
chapa=request.form("chapa").item(a)
sql="select f.chapa, f.nome, f.dataadmissao, f.sexo, s.codevento, coddoc, curso, f.secao " & _
"from dc_professor f inner join corporerm.dbo.pfsalcmp s on f.chapa collate database_default=s.chapa " & _
"left join g2cursoeve c on c.sal=s.codevento collate database_default " & _
"where f.chapa='" & chapa & "' order by coddoc, nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
sufixo=" no "
do while not rs.eof
if rs("sexo")="F" then s1="a" else s1="o"
if rs("sexo")="F" then s2="a" else s2=""
if rs("curso")="" or isnull(rs("curso")) then cursof=rs("secao") else cursof=rs("curso")
sql2="insert into ttavalprof (sessao, chapa, coddoc) select '" & session("usuariomaster") & "','" & rs("chapa") & "','" & rs("coddoc") & "'"
conexao.execute sql2
%>
<table border="0" cellpadding="1" cellspacing="0" width="690" height=990 style="border-collapse: collapse">
<tr><td valign=top>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class=campo align="left"><img src="../images/logo_centro_universitario_unifieo_big.jpg" width="150"  border="0" alt=""></td>
	<td class="campop" align="center"><i><b>AVALIAÇÃO DE PROFESSOR RECEM-CONTRATADO</td></tr>
<tr><td colspan=2 class="campop" align="center">Relatório de Acompanhamento Funcional</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=10></td></tr>
<tr>
	<td class="campop" valign=top height=40><font style="font-size:7pt"><i><b>Docente Avaliado:</font></b></i><br>
	&nbsp;<%=rs("chapa")%>&nbsp;&nbsp;<%=rs("nome")%><td>
	<td class=campo rowspan=3 align="right">
	<img border="0" src="../func_foto.asp?chapa=<%=rs("chapa")%>"  height="120">
	</td>
</tr>
<tr>
	<td class="campop" valign=top height=40><font style="font-size:7pt"><i><b>Curso Principal:</font></b></i><br>
	&nbsp;<%=cursof%><td>
</tr>
<tr>
	<td class="campop" valign=top height=40><font style="font-size:7pt"><i><b>Data de admissão:</font></b></i><br>
	&nbsp;<%=rs("dataadmissao")%><td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=5></td></tr>
<tr><td class="campop">Prezado(a) Coordenador(a)<br>
	Preencha a avaliação com base no desempenho d<%=s1%> professor<%=s2%> e ao término devolva ao Deptº de Recursos Humanos.
	<br>Orientação:
	<br>a) Avalie o quesito e considere o valor indicado em cada um.
	<br>b) Some o total de cada coluna e apure a média.
	</td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" colspan=9 height=10></td></tr>
<tr><td class=fundop colspan=4><b>Quesitos a serem observados</td>
	<td class=fundop width="45" align="center"><b>Ótimo  <br>(5)</td>
	<td class=fundop width="45" align="center"><b>Bom    <br>(4)</td>
	<td class=fundop width="45" align="center"><b>Regular<br>(3)</td>
	<td class=fundop width="45" align="center"><b>Fraco  <br>(2)</td>
	<td class=fundop width="45" align="center"><b>Ruim   <br>(1)</td>
</tr>
<%
for b=1 to 14
col2=100
%>
<tr><td class="campop" align="center" rowspan=1 width=25 style="border-right:1px solid #000000;border-top:1px dotted #000000"><b><%=b%></td>
	<td class=campo width=5></td>
	<td class="campop" height=30 style="border-top:1px dotted #000000"><b> <%=perg(b)%> </td>

	<td class=campo width=5></td><%col=5%>
	<%for c=1 to 5%>
	<td class=campo widht="35" align="center" style="border-top:1px dotted #000000"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<%next%>
</tr>
<%
next
%>
<tr><td class="campor" colspan=9 height=5></td></tr>
<tr><td class=fundop colspan=4 height=30 align="center"><b>Sub-totais</td>
	<td class=fundop width="40" style="border-left:1px solid #000000" align="center">&nbsp;</td>
	<td class=fundop width="40" style="border-left:1px solid #000000" align="center">&nbsp;</td>
	<td class=fundop width="40" style="border-left:1px solid #000000" align="center">&nbsp;</td>
	<td class=fundop width="40" style="border-left:1px solid #000000" align="center">&nbsp;</td>
	<td class=fundop width="40" style="border-left:1px solid #000000;border-right:1px solid #000000" align="center">&nbsp;</td>
</tr>
</table>
<br>

<div align="center">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campop" style="border-bottom:1px dotted" width="150">Somatório colunas</td>
	<td class="campop" width="25">&nbsp;</td>
	<td class="campop" style="border-bottom:1px dotted" width="100">Média final</td>
</tr>
<tr>
	<td class="campop" style="border-bottom:1px dotted #000000" align="left">=</td>
	<td class="campop" align="center"> ÷ 14 </td>
	<td class=cmapop style="border-bottom:1px dotted #000000" align="left">=</td>
</tr>
</table>
</div>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=10></td></tr>
<tr>
	<td width=345 height=50 class="campor" valign=top style="border-top:1px solid;border-left:1px solid;border-bottom:1px solid">&nbsp;Avaliador/Coordenador</td>
	<td width=230 class="campor" valign=top style="border-top:1px solid;border-left:1px solid;border-bottom:1px solid">&nbsp;Assinatura</td>
	<td width=115 class="campor" valign=top style="border:1px solid">&nbsp;Data</td>
</td>
</table>


<br>
<br>
<table border="1" bordercolor="#000000" cellpadding="3" cellspacing="0" width=100% style="border-collapse: collapse">
<tr>
	<td class=fundo nowrap>Média Final:</td>
	<td class=fundo width="100" align="center">1</td>
	<td class=fundo width="100" align="center">2</td>
	<td class=fundo width="100" align="center">3</td>
	<td class=fundo width="100" align="center">4</td>
	<td class=fundo width="100" align="center">5</td>
</tr>
<tr>
	<td class=fundo>Orientação</td>
	<td class=fundo align="center" colspan=2 width="200" >
		Indique os pontos a serem desenvolvidos, caso não demonstre melhora até o final do período
		letivo, sugerir a dispensa</td>
	<td class=fundo align="center" width="100" >
		Aprovado</td>
	<td class=fundo align="center" colspan=2 width="200" >
		Aprovado / Indicar para outros coordenadores de curso</td>
</tr>

</table>

</td></tr>
<tr><td class=campo height=100%></td></tr>
</table>
<%
rs.movenext
response.write "<DIV style=""page-break-after:always""></DIV>"
loop
rs.close

'****
next
'****
response.write "</table>"
%>
<%
'******************* cartas para acompanhar

sql3="select a.coddoc, curso, coordenador from ttavalprof a, g2cursoeve c where sessao='" & session("usuariomaster") & "' and c.coddoc=a.coddoc group by a.coddoc, curso, coordenador "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
<table border="0" cellpadding="1" cellspacing="0" width="690" height=990 style="border-collapse: collapse">
<tr><td valign=top>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
	<tr><td class="campop"><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width=225></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><b>Of. RH-<%=rs("coddoc")%></b></td></tr>
	<tr><td class="campop" align="right">
	<input type="text" name="txt1" class="form_input" size="29" value="Osasco, <%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %>" style="font-size:10pt">
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><input type="text" name="txt0" class="form_input" size="5" value=""><br>
	<input type="text" name="txt1" class="form_input" size="60" value="Ilmo(a) Sr(a). <%=rs("coordenador")%>" style="font-size:10pt"><br>
	<input type="text" name="txt2" class="form_input" size="60" value="<%=rs("curso")%>" style="font-size:10pt"><br>
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Em atendimento a vossa solicitação, foram contratados para prestarem serviços a esse curso, os seguintes professores:

	<div align="center">
	<table border="1" cellpadding="3" cellspacing="0" width=80% style="border-collapse: collapse">
	<tr><td class=titulop>Nome</td>
		<td class=titulop align="center">Admissão</td>
	</tr>
<%
sql4="select a.chapa, f.nome, f.dataadmissao from ttavalprof a, dc_professor f where a.chapa=f.chapa and a.coddoc='" & rs("coddoc") & "' and a.sessao='" & session("usuariomaster") & "'"
rs2.Open sql4, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
	<tr><td class="campop"><%=rs2("nome")%></td>
		<td class="campop" align="center"><%=rs2("dataadmissao")%></td>
	</tr>
<%
rs2.movenext
loop
rs2.close
%>	
	</table></div>
	<p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Solicitamos que V.Sa. se manifeste por escrito, na avaliação anexa sobre o desempenho dos referidos professores, 
	informando-nos, no prazo de 10 (dez) dias, se os mesmos atendem as exigências do cargo e preenche 
	os requisitos indispensáveis. Em resumo, se amolda aos padrões da FIEO.</td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Contando com a costumeira 
	colaboração de V.Sa. apresentamos nossas cordiais saudações.</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><p align="center" style="line-height: 25px">
	<input type="text" name="txt1" class="form_input" size="60" value="MARIA CÉLIA SOARES SOARES HUNGRIA" style="font-size:10pt;text-align:center"><br>
	<input type="text" name="txt2" class="form_input" size="30" value="Pró-Reitoria Acadêmica" style="font-size:10pt;text-align:center">
	</td></tr>
	<tr><td class="campop"></td></tr>
</table>
	
</td></tr>
<tr><td class=campo height=100%></td></tr>
</table>
<%
rs.movenext
response.write "<DIV style=""page-break-after:always""></DIV>"
loop
rs.close

end if 'tipo T/C
%>

<%
if tipo="T" or tipo="A" then

for a=1 to chapas
'****
chapa=request.form("chapa").item(a)
sql="select f.chapa, f.nome, f.dataadmissao, f.sexo, s.codevento, coddoc, curso, f.secao " & _
"from dc_professor f inner join corporerm.dbo.pfsalcmp s on f.chapa collate database_default=s.chapa " & _
"left join g2cursoeve c on c.sal=s.codevento collate database_default " & _
"where f.chapa='" & chapa & "' order by coddoc, nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof
sql2="insert into ttavalprof (sessao, chapa, coddoc) select '" & session("usuariomaster") & "','" & rs("chapa") & "','" & rs("coddoc") & "'"
conexao.execute sql2
rs.movenext
loop
rs.close
next

dim aperg(14)
'******************* fichas de auto-avaliação
aperg(1)="Possui boa apresentação pessoal"
aperg(2)="Teve fácil adaptação e respeito às normas da Instituição"
aperg(3)="Costuma chegar e sair da sala de aula no horário correto"
aperg(4)="Apresenta assiduidade, comparece para ministrar aula, raramente falta"
aperg(5)="Cumpre o plano de ensino e segue as bibliografias indicadas"
aperg(6)="Apresenta conteúdos atualizados e utiliza sempre que necessário recursos áudio-visuais em suas aulas"
aperg(7)="Apresenta didática e conhecimentos necessários para transmissão dos conteúdos aos alunos"
aperg(8)="Cumpre prazos para entrega e publicação das notas de avaliações e/ou trabalhos"
aperg(9)="Relaciona-se de forma positiva com os alunos (disciplina, respeito e confiança)"
aperg(10)="Os alunos estão satisfeitos com sua forma de lecionar. Não existe registros de queixas na coordenação e ouvidoria"
aperg(11)="Tem facilidade de relacionamento com os colegas e coordenador"
aperg(12)="Está disposto a colaborar com os demais, tem iniciativa e entusiasmo para trabalhos coletivos"
aperg(13)="Participa nas reuniões de trabalho, demonstrando interesse e envolvimento"
aperg(14)="Considera-se apto para lecionar em outro curso ou disciplina"

sql3="select f.chapa, f.nome, f.dataadmissao, f.sexo, f.secao " & _
"from dc_professor f " & _
"where f.chapa in (select distinct chapa from ttavalprof where sessao='" & session("usuariomaster") & "') order by secao, nome "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
<table border="0" cellpadding="1" cellspacing="0" width="690" height=990 style="border-collapse: collapse">
<tr><td valign=top>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class=campo align="left"><img src="../images/logo_centro_universitario_unifieo_big.jpg" width="150"  border="0" alt=""></td>
	<td class="campop" align="center"><i><b>AUTO-AVALIAÇÃO DE PROFESSOR RECEM-CONTRATADO</td></tr>
<tr><td colspan=2 class="campop" align="center">Relatório de Acompanhamento Funcional</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=10></td></tr>
<tr>
	<td class="campop" valign=top height=40><font style="font-size:7pt"><i><b>Docente Avaliado:</font></b></i><br>
	&nbsp;<%=rs("chapa")%>&nbsp;&nbsp;<%=rs("nome")%><td>
	<td class=campo rowspan=3 align="right">
	<img border="0" src="../func_foto.asp?chapa=<%=rs("chapa")%>"  height="120">
	</td>
</tr>
<tr>
	<td class="campop" valign=top height=40><font style="font-size:7pt"><i><b>Curso Principal:</font></b></i><br>
	&nbsp;<%=rs("secao")%><td>
</tr>
<tr>
	<td class="campop" valign=top height=40><font style="font-size:7pt"><i><b>Data de admissão:</font></b></i><br>
	&nbsp;<%=rs("dataadmissao")%><td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=5></td></tr>
<tr><td class="campop">Prezado(a) Professor(a)<br>
	Preencha a avaliação com base no seu desempenho e ao término devolva ao Deptº de Recursos Humanos.
	<br>Orientação:
	<br>a) Avalie o quesito e considere o valor indicado em cada um.
	<br>b) Some o total de cada coluna e apure a média.
	</td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" colspan=9 height=10></td></tr>
<tr><td class=fundop colspan=4><b>Quesitos a serem observados</td>
	<td class=fundop width="45" align="center"><b>Ótimo  <br>(5)</td>
	<td class=fundop width="45" align="center"><b>Bom    <br>(4)</td>
	<td class=fundop width="45" align="center"><b>Regular<br>(3)</td>
	<td class=fundop width="45" align="center"><b>Fraco  <br>(2)</td>
	<td class=fundop width="45" align="center"><b>Ruim   <br>(1)</td>
</tr>
<%
for b=1 to 14
col2=100
%>
<tr><td class="campop" align="center" rowspan=1 width=25 style="border-right:1px solid #000000;border-top:1px dotted #000000"><b><%=b%></td>
	<td class=campo width=5></td>
	<td class="campop" height=30 style="border-top:1px dotted #000000"><b> <%=aperg(b)%> </td>

	<td class=campo width=5></td><%col=5%>
	<%for c=1 to 5%>
	<td class=campo widht="35" align="center" style="border-top:1px dotted #000000"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<%next%>
</tr>
<%
next
%>
<tr><td class="campor" colspan=9 height=5></td></tr>
<tr><td class=fundop colspan=4 height=30 align="center"><b>Sub-totais</td>
	<td class=fundop width="40" style="border-left:1px solid #000000" align="center">&nbsp;</td>
	<td class=fundop width="40" style="border-left:1px solid #000000" align="center">&nbsp;</td>
	<td class=fundop width="40" style="border-left:1px solid #000000" align="center">&nbsp;</td>
	<td class=fundop width="40" style="border-left:1px solid #000000" align="center">&nbsp;</td>
	<td class=fundop width="40" style="border-left:1px solid #000000;border-right:1px solid #000000" align="center">&nbsp;</td>
</tr>
</table>
<br>

<div align="center">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campop" style="border-bottom:1px dotted" width="150">Somatório colunas</td>
	<td class="campop" width="25">&nbsp;</td>
	<td class="campop" style="border-bottom:1px dotted" width="100">Média final</td>
</tr>
<tr>
	<td class="campop" style="border-bottom:1px dotted #000000" align="left">=</td>
	<td class="campop" align="center"> ÷ 14 </td>
	<td class=cmapop style="border-bottom:1px dotted #000000" align="left">=</td>
</tr>
</table>
</div>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=10></td></tr>
<tr>
	<td width=230 class="campor" valign=top style="border-top:1px solid;border-left:1px solid;border-bottom:1px solid">&nbsp;Assinatura do Professor</td>
	<td width=115 class="campor" valign=top style="border:1px solid">&nbsp;Data</td>
	<td width=345 height=50 class=fundor valign=top style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid">&nbsp;</td>
</td>
</table>

<br>

</td></tr>
<tr><td class=campo height=100%></td></tr>
</table>

<%
rs.movenext
response.write "<DIV style=""page-break-after:always""></DIV>"
loop
rs.close
end if 'tipo T/A

%>

<%
end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>