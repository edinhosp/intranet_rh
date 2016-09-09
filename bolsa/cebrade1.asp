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
<title>Convênio CEBRADE</title>
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
<p class=titulo style="margin-top: 0; margin-bottom: 0">Convênio CEBRADE (Centro Brasileiro de Desenvolvimento do Ensino Superior)</p>
<p>
<input type="radio" name="opcao" value="2">Anexo II - Requerimento de Adesão ao Termo de Convênio
<br>
<input type="radio" name="opcao" value="3">Anexo III - Termo de Convênio PAET de Concessão de Bolsas de Estudo
<br>
<input type="radio" name="opcao" value="4">Anexo IV - Termo Aditivo de Inclusão de Aluno no Convênio PAET de Concessão de Bolsas de Estudo
<br>
Informações Complementares:<br>
<br>Representante: <input type="text" name="representante" size="45" value="JOSÉ CASSIO SOARES HUNGRIA">
<br>RG: <input type="text" name="RG" size="15" value="1.409.223"> - SSP/<input type="text" name="UF" size="2" value="SP">
<br>CPF: <input type="text" name="CPF" size="15" value="037.195.298-00">
<br>
<br><input type="submit" value="Visualizar">
</form>
<%
end if 'formulário inicial

'<!-- ************************* OPCAO 2 ************************* -->

if request.form("opcao")="2" then
%>
<!-- tabela quadro de página -->
<div align="right">
<table border="0" width="650" height="1000" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr><td class="campo" valign="top">
<br><br><br>
<p align="center"><b><i>REQUERIMENTO DE ADESÃO AO TERMO DE CONVÊNIO</i></b></p>
<br><br><br>
<p>Ao:<br>
Centro Brasileiro de Desenvolvimento do Ensino Superior - CEBRADE</p>
<br><br><br>
<p align="justify">A FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO, representada neste ato por seu representante legal Sr. <%=request.form("representante")%>, 
portador do RG nº <%=request.form("RG")%> - SSP/<%=request.form("UF")%> e do CPF nº <%=request.form("CPF")%>, com sede
na Avenida Franz Voegeli, 300 - Vila Yara - Osasco - SP, vem, por meio da presente, nos termos do que estabelece a
Convenção Coletiva de Trabalho e Regulamento do Programa de Capacitação, requerer a adesão ao Termo de Convênio PAET 
de Concessão de Bolsas de Estudo, cujos alunos participantes seguem abaixo:</p>

<!-- -->
<table border="1" bordercolor="#000000" width="630"  cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campo" valign="middle" align="center">Nome do Aluno</td>
	<td class="campo" valign="middle" align="center">Matrícula</td>
	<td class="campo" valign="middle" align="center">Curso</td>
	<td class="campo" valign="middle" align="center">Série</td>
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
	<td class="campo" valign="middle" align="center">Matrícula</td>
	<td class="campo" valign="middle" align="center">Curso</td>
	<td class="campo" valign="middle" align="center">Série</td>
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
<!-- fim tabela quadro de página -->

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
<p align="center"><b><i>TERMO DE CONVÊNIO PAET DE CONCESSÃO DE BOLSAS DE ESTUDO</i></b></p>
<br>

<p align="justify">Pelo presente instrumento, de um lado CENTRO BRASILEIRO DE DESENVOLVIMENTO DO
ENSINO SUPERIOR – CEBRADE, pessoa jurídica de direito privado, sem fins lucrativos, inscrita no
CNPJ sob n.º 05.578.073/0001-89, domiciliada na Rua Cipriano Barata, 2431 – Ipiranga – São
Paulo – SP, representado neste ato pelo Sr. Gabriel Mário Rodrigues, doravante denominado
CEBRADE e de outro lado a FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO, entidade doravante denominada abreviadamente
INSTITUIÇÃO, representada neste ato por seu representante legal Sr. <%=request.form("representante")%>, 
portador do RG n.° <%=request.form("RG")%> - SSP/<%=request.form("UF")%> e do CPF n° <%=request.form("CPF")%>, 
com sede na Avenida Franz Voegeli, 300 - Vila Yara - Osasco - SP, considerando a necessidade
de implementar um sistema de concessão de bolsas aos dependentes de professores e auxiliares
da educação superior mediante o desenvolvimento do Programa de Amparo Educativo Temporário
– PAET, que priorize o desenvolvimento, integração e acesso à Educação Superior no Estado São
Paulo, resolvem celebrar o presente convênio de cooperação, e de acordo com as cláusulas e
condições a seguir:</b>

<p style="margin-bottom:0px;margin-top:10px"><b>DO OBJETO</b></p>
<p style="margin-bottom:0px;margin-top:0px"><b>CLÁUSULA PRIMEIRA</b></p>
<p style="margin-bottom:0px;margin-top:0px" align="justify">O presente Convênio tem por objeto estabelecer, em regime de cooperação mútua entre os
partícipes, o desenvolvimento da educação superior no país mediante a concessão de bolsas de
estudo aos dependentes legais dos empregados das instituições de ensino superior participantes do
presente convênio.

<p style="margin-bottom:0px;margin-top:10px"><b>DAS CONDIÇÕES GERAIS
<p style="margin-bottom:0px;margin-top:0px"><b>CLÁUSULA SEGUNDA
<p style="margin-bottom:0px;margin-top:0px" align="justify">Fica estabelecido entre as partes que o CEBRADE – Centro Brasileiro de Desenvolvimento do
Ensino Superior – que possui como um dos seus objetivos, desenvolvimento do Programa de
Amparo Educativo Temporário – PAET, concedendo bolsas de estudo em Instituições Privadas de
Ensino Superior concederá aos filhos ou dependentes legais do empregado o direito de usufruir as
gratuidades integrais do PAET, sem qualquer ônus, nos cursos de graduação e sequencial
existentes e administrados pela INSTITUIÇÃO para a qual o empregado trabalha, observado o
disposto neste instrumento.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO PRIMEIRO</b>. A INSTITUIÇÃO deverá disponibilizar ao CEBRADE, mediante
requerimento, bolsas de estudo em número suficiente para o atendimento da concessão das
gratuidades integrais do PAET nas Instituições de Ensino Superior por ela mantida, para filhos ou
dependentes legais dos seus empregados, observada a limitação estabelecida na cláusula de
bolsas de estudo.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO SEGUNDO</b>. Para a concessão das gratuidades integrais aos filhos e dependentes
legais do empregado, o CEBRADE não poderá fazer qualquer outra exigência a não ser o
comprovante de aprovação no processo seletivo da INSTITUIÇÃO empregadora e a observância
dos preceitos estabelecidos neste instrumento.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO TERCEIRO</b>. Terão direito a requerer e obter do CEBRADE a concessão de bolsas
integrais de estudo, os dependentes legais do empregado reconhecidos pela Legislação do Imposto
de Renda, ou que estejam sob a sua guarda judicial e vivam sob sua dependência econômica,
devidamente comprovada.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO QUARTO</b>. Os filhos do empregado terão direito a obter do CEBRADE concessão de
bolsas de estudo integrais, desde que, na data de efetivação da matrícula no curso superior, não
tenham 25 (vinte e cinco anos) completos ou mais.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO QUINTO</b>. As bolsas de estudo são válidas para cursos de graduação e sequenciais e
a INSTITUIÇÃO está obrigada a conceder, no máximo, duas bolsas de estudo por empregado.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO SEXTO</b>. O beneficiário bolsista, concluinte de curso de graduação não poderá obter
nova concessão de gratuidade na mesma instituição.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO SÉTIMO</b>. As bolsas de estudo serão mantidas aos dependeste quando o empregado
estiver licenciado para tratamento de saúde ou em gozo de licença mediante anuência da
INSTITUIÇÃO, excetuado quando o empregado tiver licenciado por “Licença sem Remuneração”.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO OITAVO</b>. No caso de falecimento do empregado, os dependentes legais que já se
encontrarem estudando na INSTITUIÇÃO continuarão a gozar das bolsas de estudo até o final do
curso.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO NONO</b>. No caso de dispensa sem justa causa do empregado durante o período
letivo, ficam garantidas até o final do período letivo, as bolsas de estudo já existentes.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO DÉCIMO</b>. Os bolsistas que forem reprovados no período letivo perderão o direito à
bolsa de estudo, voltando a gozar do benefício quando lograrem aprovação no referido período. As
disciplinas cursadas em regime de dependência serão de total responsabilidade do bolsista,
arcando o mesmo com o seu custo.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO DÉCIMO PRIMEIRO</b>. Além dos casos previstos nesta cláusula, a INSTITUIÇÃO
poderá fornecer outras bolsas de estudos, cujas condições serão objeto de termo aditivo a ser
firmado entre a INSTITUIÇÃO e o CEBRADE, nos termos do ANEXO IV.

<p style="margin-bottom:0px;margin-top:10px"><b>DA COMISSÃO DE ACOMPANHAMENTO DO CONVÊNIO
<p style="margin-bottom:0px;margin-top:0px"><b>CLÁUSULA TERCEIRA
<p style="margin-bottom:0px;margin-top:0px" align="justify">O SEMESP e a Federação representante da categoria profissiona fiscalizará o CEBRADE na
gestão do Programa de Amparo Educativo Temporário para os filhos e dependentes legais dos
empregados nas instituições de ensino pertencentes a sua categoria representativa.
<p style="margin-bottom:0px;margin-top:5px" align="justify"><b>PARÁGRAFO ÚNICO</b>. Os convenentes desde já expressam concordância quanto à fiscalização,
bem como se comprometem a fornecer todos os documentos que lhe forem solicitados para
comprovar o cumprimento das obrigações ora assumidas.

<p style="margin-bottom:0px;margin-top:10px"><b>DO PRAZO
<p style="margin-bottom:0px;margin-top:0px"><b>CLÁUSULA QUARTA
<p style="margin-bottom:0px;margin-top:0px" align="justify">O presente Convênio vigorará até 29 de fevereiro de 2013, tendo como termo inicial a data de sua
assinatura, podendo ser renovado no interesse dos partícipes por novos prazos.

<p style="margin-bottom:0px;margin-top:10px"><b>DO DESCUMPRIMENTO DAS OBRIGAÇÕES
<p style="margin-bottom:0px;margin-top:0px"><b>CLÁUSULA QUINTA
<p style="margin-bottom:0px;margin-top:0px" align="justify">O descumprimento pelos convenentes dos compromissos assumidos neste convênio ensejará a
rescisão do presente instrumento e a aplicação das penalidades previstas na Lei.

<p style="margin-bottom:0px;margin-top:10px"><b>CONFIDENCIALIDADE
<p style="margin-bottom:0px;margin-top:0px"><b>CLÁUSULA SEXTA
<p style="margin-bottom:0px;margin-top:0px" align="justify">Comprometem-se as partes a proteger as informações confidenciais, no caso do presente
instrumento dados pessoais e qualquer outro informado na “Solicitação de bolsa de estudo”, sob
pena de responder pelos danos causados, sem prejuízo de indenização e outras medidas cabíveis.

<p style="margin-bottom:0px;margin-top:10px"><b>DO FORO
<p style="margin-bottom:0px;margin-top:0px" align="justify">Em caso de controvérsias, oriundas do presente convênio, as partes, desde já, elegem o Foro da
Capital de São Paulo, por mais privilegiado que outro seja.
<p style="margin-bottom:0px;margin-top:0px"><b>CLÁUSULA SÉTIMA
<p style="margin-bottom:0px;margin-top:0px" align="justify">E, por estarem os convenentes certos e acordados quanto às cláusulas e condições deste
convênio, firmam o presente termo em 2 (duas) vias de igual teor e para um só efeito na presença
das testemunhas abaixo assinadas e qualificadas.

<p style="margin-bottom:0px;margin-top:10px">
<br>São Paulo, <%=dataextenso%>.<br>
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
<!-- tabela quadro de página -->
<div align="right">
<table border="0" width="650" height="1000" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr><td class="campo" valign="top">
<br><br><br>
<p align="center"><b><i>TERMO ADITIVO DE INCLUSÃO DE ALUNO NO CONVÊNIO PAET<BR>DE CONCESSÃO DE BOLSAS DE ESTUDO</i></b></p>
<br><br><br>
<p>Ao:<br>
Centro Brasileiro de Desenvolvimento do Ensino Superior - CEBRADE</p>
<br><br><br>
<p align="justify">A FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO, representada neste ato por seu representante legal Sr. <%=request.form("representante")%>, 
portador do RG nº <%=request.form("RG")%> - SSP/<%=request.form("UF")%> e do CPF nº <%=request.form("CPF")%>, com sede
na Avenida Franz Voegeli, 300 - Vila Yara - Osasco - SP, vem, por meio da presente, nos termos do que estabelece a
Convenção Coletiva de Trabalho e Regulamento da Cláusula de Bolsa de Estudos, solicitar a inclusão dos alunos abaixo
indicados no Termo de Convênio PAET de Concessão de Bolsas de Estudos:</p>

<!-- -->
<table border="1" bordercolor="#000000" width="630"  cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campo" valign="middle" align="center">Nome do Aluno</td>
	<td class="campo" valign="middle" align="center">Matrícula</td>
	<td class="campo" valign="middle" align="center">Curso</td>
	<td class="campo" valign="middle" align="center">Série</td>
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
	<td class="campo" valign="middle" align="center">Matrícula</td>
	<td class="campo" valign="middle" align="center">Curso</td>
	<td class="campo" valign="middle" align="center">Série</td>
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
<!-- fim tabela quadro de página -->

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