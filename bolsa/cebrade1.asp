<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a67")="N" or session("a67")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Conv�nio CEBRADE</title>
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
registros=Session("RegistrosPorPagina")
registros=250
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form="" then
%>
<form method="POST" name="form" action="cebrade1.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Conv�nio CEBRADE (Centro Brasileiro de Desenvolvimento do Ensino Superior)</p>
<p>
<input type="radio" name="opcao" value="2">Anexo II - Requerimento de Ades�o ao Termo de Conv�nio
<br>
<input type="radio" name="opcao" value="3">Anexo III - Termo de Conv�nio PAET de Concess�o de Bolsas de Estudo
<br>
<input type="radio" name="opcao" value="4">Anexo IV - Termo Aditivo de Inclus�o de Aluno no Conv�nio PAET de Concess�o de Bolsas de Estudo
<br>
Informa��es Complementares:<br>
<br>Representante: <input type="text" name="representante" size="45" value="JOS� CASSIO SOARES HUNGRIA">
<br>RG: <input type="text" name="RG" size="15" value="1.409.223"> - SSP/<input type="text" name="UF" size="2" value="SP">
<br>CPF: <input type="text" name="CPF" size="15" value="037.195.298-00">
<br>
<br><input type="submit" value="Visualizar">
</form>
<%
end if 'formul�rio inicial

'<!-- ************************* OPCAO 2 ************************* -->

if request.form("opcao")="2" then
%>
<!-- tabela quadro de p�gina -->
<div align="right">
<table border="0" width="650" height="1000" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr><td class="campo" valign="top">
<br><br><br>
<p align="center"><b><i>REQUERIMENTO DE ADES�O AO TERMO DE CONV�NIO</i></b></p>
<br><br><br>
<p>Ao:<br>
Centro Brasileiro de Desenvolvimento do Ensino Superior - CEBRADE</p>
<br><br><br>
<p align="justify">A FUNDA��O INSTITUTO DE ENSINO PARA OSASCO, representada neste ato por seu representante legal Sr. <%=request.form("representante")%>, 
portador do RG n� <%=request.form("RG")%> - SSP/<%=request.form("UF")%> e do CPF n� <%=request.form("CPF")%>, com sede
na Avenida Franz Voegeli, 300 - Vila Yara - Osasco - SP, vem, por meio da presente, nos termos do que estabelece a
Conven��o Coletiva de Trabalho e Regulamento do Programa de Capacita��o, requerer a ades�o ao Termo de Conv�nio PAET 
de Concess�o de Bolsas de Estudo, cujos alunos participantes seguem abaixo:</p>

<!-- -->
<table border="1" bordercolor="#000000" width="630"  cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campo" valign="middle" align="center">Nome do Aluno</td>
	<td class="campo" valign="middle" align="center">Matr�cula</td>
	<td class="campo" valign="middle" align="center">Curso</td>
	<td class="campo" valign="middle" align="center">S�rie</td>
	<td class="campor" valign="middle" align="center">Porcentagem de<br>bolsa concedida</td>
</tr>
<%
sql="declare @ano as datetime " & _
"set @ano=convert(datetime,GETDATE()) " & _
"SELECT distinct b.chapa, b.matricula, s.descricao AS situacao, t.descricao AS tipo, b.nome_bolsista " & _
", ano_letivo, b.curso, m.periodo, p.HABILITACAO " & _
"FROM ((bolsistas b INNER JOIN bolsistas_lanc l ON b.id_bolsa=l.id_bolsa) " & _
"INNER JOIN bolsistas_situacao s ON l.situacao=s.id_sit) " & _
"INNER JOIN bolsistas_tipo t ON b.tp_bolsa=t.id_tp " & _
"left join corporerm.dbo.UMATRICPL m on m.MATALUNO collate database_default=b.matricula and m.PERLETIVO collate database_default=l.ano_letivo " & _
"left join corporerm.dbo.UPERIODOS p on p.codcur=m.CODCUR and p.codper=m.codper " & _
"WHERE b.tp_bolsa In ('2') AND @ano between l.renovacao and l.validade and id_sit not in ('I') " & _
"and m.STATUS not in (53) " & _
"ORDER BY nome_bolsista"
rs.CursorLocation=3
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalb=rs.recordcount
do while not rs.eof

if rs.absoluteposition>25 and pulou=0 then
%>
</table>
<DIV style="page-break-after:always"></DIV>
<table border="1" bordercolor="#000000" width="630"  cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campo" valign="middle" align="center">Nome do Aluno</td>
	<td class="campo" valign="middle" align="center">Matr�cula</td>
	<td class="campo" valign="middle" align="center">Curso</td>
	<td class="campo" valign="middle" align="center">S�rie</td>
	<td class="campor" valign="middle" align="center">Porcentagem de<br>bolsa concedida</td>
</tr>
<%
	pulou=1
end if

%>
<tr>
	<td class="campo" height="25" valign="middle" align="left"><%=rs("nome_Bolsista")%></td>
	<td class="campo" valign="middle" align="center"><%=rs("matricula")%></td>
	<td class="campor" valign="middle" align="left"><%=rs("habilitacao")%></td>
	<td class="campo" valign="middle" align="center"><%=rs("periodo")%></td>
	<td class="campo" valign="middle" align="center"><input type="text" class="form_input" size="6" value="100%"></td>
</tr>
<%
rs.movenext
loop
rs.close
%>
<tr><td class="campo" colspan="5">Total de bolsistas informados: <%=totalb%></td></tr>
</table>

<!-- -->
<br><br><br>
_____________________________________________________<br>
(Assinatura do representante legal da Mantenedora)

</td></tr>
</table>
</div>
<!-- fim tabela quadro de p�gina -->

<%
end if

'<!-- ************************* OPCAO 3 ************************* -->

if request.form("opcao")="3" then
	dataextenso=day(now()) & " de " & monthname(month(now())) & " de " & year(now())

%>
<div align="right">
<table border="0" width="650" height="1000" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr><td class="campo" valign="top">
<br>
<p align="center"><b><i>TERMO DE CONV�NIO PAET DE CONCESS�O DE BOLSAS DE ESTUDO</i></b></p>
<br>

<p align="justify">Pelo presente instrumento, de um lado CENTRO BRASILEIRO DE DESENVOLVIMENTO DO
ENSINO SUPERIOR � CEBRADE, pessoa jur�dica de direito privado, sem fins lucrativos, inscrita no
CNPJ sob n.� 05.578.073/0001-89, domiciliada na Rua Cipriano Barata, 2431 � Ipiranga � S�o
Paulo � SP, representado neste ato pelo Sr. Gabriel M�rio Rodrigues, doravante denominado
CEBRADE e de outro lado a FUNDA��O INSTITUTO DE ENSINO PARA OSASCO, entidade doravante denominada abreviadamente
INSTITUI��O, representada neste ato por seu representante legal Sr. <%=request.form("representante")%>, 
portador do RG n.� <%=request.form("RG")%> - SSP/<%=request.form("UF")%> e do CPF n� <%=request.form("CPF")%>, 
com sede na Avenida Franz Voegeli, 300 - Vila Yara - Osasco - SP, considerando a necessidade
de implementar um sistema de concess�o de bolsas aos dependentes de professores e auxiliares
da educa��o superior mediante o desenvolvimento do Programa de Amparo Educativo Tempor�rio
� PAET, que priorize o desenvolvimento, integra��o e acesso � Educa��o Superior no Estado S�o
Paulo, resolvem celebrar o presente conv�nio de coopera��o, e de acordo com as cl�usulas e
condi��es a seguir:</b>

<p style="margin-bottom:0px;margin-top:10px"><b>DO OBJETO</b></p>
<p style="margin-bottom:0px;margin-top:0px"><b>CL�USULA PRIMEIRA</b></p>
<p style="margin-bottom:0px;margin-top:0px" align="justify">O presente Conv�nio tem por objeto estabelecer, em regime de coopera��o m�tua entre os
part�cipes, o desenvolvimento da educa��o superior no pa�s mediante a concess�o de bolsas de
estudo aos dependentes legais dos empregados das institui��es de ensino superior participantes do
presente conv�nio.

<p style="margin-bottom:0px;margin-top:10px"><b>DAS CONDI��ES GERAIS
<p style="margin-bottom:0px;margin-top:0px"><b>CL�USULA SEGUNDA
<p style="margin-bottom:0px;margin-top:0px" align="justify">Fica estabelecido entre as partes que o CEBRADE � Centro Brasileiro de Desenvolvimento do
Ensino Superior � que possui como um dos seus objetivos, desenvolvimento do Programa de
Amparo Educativo Tempor�rio � PAET, concedendo bolsas de estudo em Institui��es Privadas de
Ensino Superior conceder� aos filhos ou dependentes legais do empregado o direito de usufruir as
gratuidades integrais do PAET, sem qualquer �nus, nos cursos de gradua��o e sequencial
existentes e administrados pela INSTITUI��O para a qual o empregado trabalha, observado o
disposto neste instrumento.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO PRIMEIRO</b>. A INSTITUI��O dever� disponibilizar ao CEBRADE, mediante
requerimento, bolsas de estudo em n�mero suficiente para o atendimento da concess�o das
gratuidades integrais do PAET nas Institui��es de Ensino Superior por ela mantida, para filhos ou
dependentes legais dos seus empregados, observada a limita��o estabelecida na cl�usula de
bolsas de estudo.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO SEGUNDO</b>. Para a concess�o das gratuidades integrais aos filhos e dependentes
legais do empregado, o CEBRADE n�o poder� fazer qualquer outra exig�ncia a n�o ser o
comprovante de aprova��o no processo seletivo da INSTITUI��O empregadora e a observ�ncia
dos preceitos estabelecidos neste instrumento.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO TERCEIRO</b>. Ter�o direito a requerer e obter do CEBRADE a concess�o de bolsas
integrais de estudo, os dependentes legais do empregado reconhecidos pela Legisla��o do Imposto
de Renda, ou que estejam sob a sua guarda judicial e vivam sob sua depend�ncia econ�mica,
devidamente comprovada.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO QUARTO</b>. Os filhos do empregado ter�o direito a obter do CEBRADE concess�o de
bolsas de estudo integrais, desde que, na data de efetiva��o da matr�cula no curso superior, n�o
tenham 25 (vinte e cinco anos) completos ou mais.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO QUINTO</b>. As bolsas de estudo s�o v�lidas para cursos de gradua��o e sequenciais e
a INSTITUI��O est� obrigada a conceder, no m�ximo, duas bolsas de estudo por empregado.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO SEXTO</b>. O benefici�rio bolsista, concluinte de curso de gradua��o n�o poder� obter
nova concess�o de gratuidade na mesma institui��o.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO S�TIMO</b>. As bolsas de estudo ser�o mantidas aos dependeste quando o empregado
estiver licenciado para tratamento de sa�de ou em gozo de licen�a mediante anu�ncia da
INSTITUI��O, excetuado quando o empregado tiver licenciado por �Licen�a sem Remunera��o�.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO OITAVO</b>. No caso de falecimento do empregado, os dependentes legais que j� se
encontrarem estudando na INSTITUI��O continuar�o a gozar das bolsas de estudo at� o final do
curso.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO NONO</b>. No caso de dispensa sem justa causa do empregado durante o per�odo
letivo, ficam garantidas at� o final do per�odo letivo, as bolsas de estudo j� existentes.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO D�CIMO</b>. Os bolsistas que forem reprovados no per�odo letivo perder�o o direito �
bolsa de estudo, voltando a gozar do benef�cio quando lograrem aprova��o no referido per�odo. As
disciplinas cursadas em regime de depend�ncia ser�o de total responsabilidade do bolsista,
arcando o mesmo com o seu custo.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO D�CIMO PRIMEIRO</b>. Al�m dos casos previstos nesta cl�usula, a INSTITUI��O
poder� fornecer outras bolsas de estudos, cujas condi��es ser�o objeto de termo aditivo a ser
firmado entre a INSTITUI��O e o CEBRADE, nos termos do ANEXO IV.

<p style="margin-bottom:0px;margin-top:10px"><b>DA COMISS�O DE ACOMPANHAMENTO DO CONV�NIO
<p style="margin-bottom:0px;margin-top:0px"><b>CL�USULA TERCEIRA
<p style="margin-bottom:0px;margin-top:0px" align="justify">O SEMESP e a Federa��o representante da categoria profissiona fiscalizar� o CEBRADE na
gest�o do Programa de Amparo Educativo Tempor�rio para os filhos e dependentes legais dos
empregados nas institui��es de ensino pertencentes a sua categoria representativa.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PAR�GRAFO �NICO</b>. Os convenentes desde j� expressam concord�ncia quanto � fiscaliza��o,
bem como se comprometem a fornecer todos os documentos que lhe forem solicitados para
comprovar o cumprimento das obriga��es ora assumidas.

<p style="margin-bottom:0px;margin-top:10px"><b>DO PRAZO
<p style="margin-bottom:0px;margin-top:0px"><b>CL�USULA QUARTA
<p style="margin-bottom:0px;margin-top:0px" align="justify">O presente Conv�nio vigorar� at� 29 de fevereiro de 2013, tendo como termo inicial a data de sua
assinatura, podendo ser renovado no interesse dos part�cipes por novos prazos.

<p style="margin-bottom:0px;margin-top:10px"><b>DO DESCUMPRIMENTO DAS OBRIGA��ES
<p style="margin-bottom:0px;margin-top:0px"><b>CL�USULA QUINTA
<p style="margin-bottom:0px;margin-top:0px" align="justify">O descumprimento pelos convenentes dos compromissos assumidos neste conv�nio ensejar� a
rescis�o do presente instrumento e a aplica��o das penalidades previstas na Lei.

<p style="margin-bottom:0px;margin-top:10px"><b>CONFIDENCIALIDADE
<p style="margin-bottom:0px;margin-top:0px"><b>CL�USULA SEXTA
<p style="margin-bottom:0px;margin-top:0px" align="justify">Comprometem-se as partes a proteger as informa��es confidenciais, no caso do presente
instrumento dados pessoais e qualquer outro informado na �Solicita��o de bolsa de estudo�, sob
pena de responder pelos danos causados, sem preju�zo de indeniza��o e outras medidas cab�veis.

<p style="margin-bottom:0px;margin-top:10px"><b>DO FORO
<p style="margin-bottom:0px;margin-top:0px" align="justify">Em caso de controv�rsias, oriundas do presente conv�nio, as partes, desde j�, elegem o Foro da
Capital de S�o Paulo, por mais privilegiado que outro seja.
<p style="margin-bottom:0px;margin-top:0px"><b>CL�USULA S�TIMA
<p style="margin-bottom:0px;margin-top:0px" align="justify">E, por estarem os convenentes certos e acordados quanto �s cl�usulas e condi��es deste
conv�nio, firmam o presente termo em 2 (duas) vias de igual teor e para um s� efeito na presen�a
das testemunhas abaixo assinadas e qualificadas.

<p style="margin-bottom:0px;margin-top:10px">
<br>S�o Paulo, <%=dataextenso%>.<br>
<br>
<br>________________________________
<br>CEBRADE
<br>
<br>_________________________________
<br>MANTENEDORA
<br>
<br>TESTEMUNHA 1: ____________________________________
<br>RG:_______________________________________________
<br>CPF: ______________________________________________
<br>
<br>TESTEMUNHA 2: ____________________________________
<br>RG:_______________________________________________
<br>CPF: ______________________________________________


</td></tr>
</table>
</div>
<%
end if

'<!-- ************************* OPCAO 4 ************************* -->

if request.form("opcao")="4" then
%>
<!-- tabela quadro de p�gina -->
<div align="right">
<table border="0" width="650" height="1000" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr><td class="campo" valign="top">
<br><br><br>
<p align="center"><b><i>TERMO ADITIVO DE INCLUS�O DE ALUNO NO CONV�NIO PAET<BR>DE CONCESS�O DE BOLSAS DE ESTUDO</i></b></p>
<br><br><br>
<p>Ao:<br>
Centro Brasileiro de Desenvolvimento do Ensino Superior - CEBRADE</p>
<br><br><br>
<p align="justify">A FUNDA��O INSTITUTO DE ENSINO PARA OSASCO, representada neste ato por seu representante legal Sr. <%=request.form("representante")%>, 
portador do RG n� <%=request.form("RG")%> - SSP/<%=request.form("UF")%> e do CPF n� <%=request.form("CPF")%>, com sede
na Avenida Franz Voegeli, 300 - Vila Yara - Osasco - SP, vem, por meio da presente, nos termos do que estabelece a
Conven��o Coletiva de Trabalho e Regulamento da Cl�usula de Bolsa de Estudos, solicitar a inclus�o dos alunos abaixo
indicados no Termo de Conv�nio PAET de Concess�o de Bolsas de Estudos:</p>

<!-- -->
<table border="1" bordercolor="#000000" width="630"  cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campo" valign="middle" align="center">Nome do Aluno</td>
	<td class="campo" valign="middle" align="center">Matr�cula</td>
	<td class="campo" valign="middle" align="center">Curso</td>
	<td class="campo" valign="middle" align="center">S�rie</td>
	<td class="campor" valign="middle" align="center">Porcentagem de<br>bolsa concedida</td>
</tr>
<%
sql="declare @ano as datetime " & _
"set @ano=convert(datetime,GETDATE()) " & _
"SELECT distinct b.chapa, b.matricula, s.descricao AS situacao, t.descricao AS tipo, b.nome_bolsista " & _
", ano_letivo, b.curso, m.periodo, p.HABILITACAO " & _
"FROM ((bolsistas b INNER JOIN bolsistas_lanc l ON b.id_bolsa=l.id_bolsa) " & _
"INNER JOIN bolsistas_situacao s ON l.situacao=s.id_sit) " & _
"INNER JOIN bolsistas_tipo t ON b.tp_bolsa=t.id_tp " & _
"left join corporerm.dbo.UMATRICPL m on m.MATALUNO collate database_default=b.matricula and m.PERLETIVO collate database_default=l.ano_letivo " & _
"left join corporerm.dbo.UPERIODOS p on p.codcur=m.CODCUR and p.codper=m.codper " & _
"WHERE b.tp_bolsa In ('2') AND @ano between l.renovacao and l.validade and id_sit not in ('I') " & _
"and m.STATUS not in (53) " & _
"ORDER BY nome_bolsista"
rs.CursorLocation=3
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalb=rs.recordcount
do while not rs.eof

if rs.absoluteposition>25 and pulou=0 then
%>
</table>
<DIV style="page-break-after:always"></DIV>
<table border="1" bordercolor="#000000" width="630"  cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campo" valign="middle" align="center">Nome do Aluno</td>
	<td class="campo" valign="middle" align="center">Matr�cula</td>
	<td class="campo" valign="middle" align="center">Curso</td>
	<td class="campo" valign="middle" align="center">S�rie</td>
	<td class="campor" valign="middle" align="center">Porcentagem de<br>bolsa concedida</td>
</tr>
<%
	pulou=1
end if

%>
<tr>
	<td class="campo" height="25" valign="middle" align="left"><%=rs("nome_Bolsista")%></td>
	<td class="campo" valign="middle" align="center"><%=rs("matricula")%></td>
	<td class="campor" valign="middle" align="left"><%=rs("habilitacao")%></td>
	<td class="campo" valign="middle" align="center"><%=rs("periodo")%></td>
	<td class="campo" valign="middle" align="center"><input type="text" class="form_input" size="6" value="100%"></td>
</tr>
<%
rs.movenext
loop
rs.close
%>
<tr><td class="campo" colspan="5">Total de bolsistas informados: <%=totalb%></td></tr>
</table>

<!-- -->
<br><br><br>
_____________________________________________________<br>
(Assinatura do representante legal da Mantenedora)

</td></tr>
</table>
</div>
<!-- fim tabela quadro de p�gina -->

<%
end if
%>

</body>
</html>
<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>