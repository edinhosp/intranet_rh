<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a45")="N" or session("a45")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pontos importantes da Conven��o Coletiva 2007</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<!-- -->
<!-- -->
<%
'dim conexao, rs, rs2
'set conexao=server.createobject ("ADODB.Connection")
'conexao.Open application("conexao")
'set rs=server.createobject ("ADODB.Recordset")
'Set rs.ActiveConnection = conexao
'sqla="SELECT dc_carga.CURSO FROM dc_carga GROUP BY dc_carga.CURSO;"
'rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<!-- auxiliares -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td valign=top class=titulo colspan=3>Pontos importantes das Conven��es Coletivas 2007</td></tr>
<tr><td valign=top class=grupo colspan=3>Auxiliares</td></tr>
<tr><td valign=top class=titulo>Cl�usula</td>
	<td valign=top class=titulo>Teor atual</td>
	<td valign=top class=titulo>Teor anterior</td></tr>

<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	13. Anota��es na Carteira de Trabalho
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo �nico � � obrigat�ria a anota��o na CTPS das mudan�as provocadas por ascens�o em plano de carreira. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo �nico - � obrigat�ria a anota��o na CTPS das mudan�as provocadas por ascens�o em plano de carreira <font style="text-decoration:line-through;font-weight:bold">ou altera��o de titula��o</font>.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	16. Bolsas de Estudo
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Todo AUXILIAR tem direito a bolsas de estudo integrais, incluindo matr�cula, para si, <b>c�njuge</b>, filhos ou dependentes legais, ambos entendidos como aqueles reconhecidos pela legisla��o do Imposto de Renda ou aqueles que estejam sob a guarda judicial do AUXILIAR e vivam sob sua depend�ncia econ�mica, devidamente comprovada.	
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo 1� - Somente ter�o direito a bolsas de estudo integrais, o(a) AUXILIAR, <font style="text-decoration:line-through;font-weight:bold">esposo(a) e companheiro(a)</font>, bem como seus filhos(as) e dependentes legais que estejam sob a guarda judicial, estes dois �ltimos desde que tenham 25 (vinte e cinco) anos ou menos na data de realiza��o do exame vestibular ou do processo seletivo que define o ingresso no curso superior.	
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	16. Bolsas de Estudo
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo s�timo � As bolsas de estudo integrais em cursos de p�s-gradua��o ou especializa��o existentes e administrados pela MANTENEDORA s�o v�lidas exclusivamente para o AUXILIAR, respeitados os crit�rios de sele��o exigidos para ingresso nos mesmos e obedecer�o �s seguintes condi��es: 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo oitavo - As bolsas de estudo integrais em cursos de p�s-gradua��o ou de especializa��o existentes e administrados pela MANTENEDORA s�o v�lidas exclusivamente para o AUXILIAR <font style="text-decoration:line-through;font-weight:bold">em �reas correlatas �quelas em que o AUXILIAR exerce a fun��o na MANTENEDORA e que visem � sua capacita��o,</font> respeitados os crit�rios de sele��o exigidos para ingresso nos mesmos e obedecer�o �s seguintes condi��es:
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	16. Bolsas de Estudo
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo onze - Quando, a crit�rio da MANTENEDORA, o AUXILIAR, em raz�o das fun��es exercidas na Institui��o se vir na conting�ncia de efetuar seus estudos, na �rea educacional indicada em outra institui��o de ensino, a MANTENEDORA arcar� com o valor integral das mensalidades do curso, incluindo matr�cula durante a vig�ncia do contrato de trabalho, respeitada a vig�ncia coletiva de trabalho.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	23. Creches	
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	� obrigat�ria a instala��o de local destinado � guarda de crian�as de at� <b>doze meses</b>, quando a unidade de ensino da MANTENEDORA mantiver contratadas, em jornada integral, pelo menos trinta funcion�rias. A manuten��o da creche poder� ser substitu�da pelo pagamento do reembolso-creche, nos termos da legisla��o em vigor (CF, 7�, XXV, Artigo 389, par�grafo 1� da CLT e Portaria MTb n� 3296 de 03.09.86), ou ainda, a celebra��o de conv�nio com uma entidade reconhecidamente id�nea. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	� obrigat�ria a instala��o de local destinado � guarda de crian�as de at� <font style="text-decoration:line-through;font-weight:bold">seis anos</font>, quando a unidade de ensino da mantiver contratadas, em jornada integral, pelo menos trinta funcion�rias<font style="text-decoration:line-through;font-weight:bold"> com idade superior a 16 anos</font>. A manuten��o da creche poder� ser substitu�da pelo pagamento do reembolso-creche, nos termos da legisla��o em vigor (CF, 7�, XXV, Artigo 389, par�grafo 1� da CLT e Portaria MTb n� 3296 de 03.09.86), ou ainda, a celebra��o de conv�nio com uma entidade reconhecidamente id�nea.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	24. Garantias ao Auxiliar em vias de Aposentadoria
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurado ao AUXILIAR que, comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria por tempo de contribui��o ou da aposentadoria por idade, a garantia de emprego durante o per�odo que faltar at� a aquisi��o do direito. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada ao AUXILIAR que, comprovadamente estiver a 24 meses ou menos da aposentadoria integral por tempo de servi�o ou da aposentadoria por idade, a garantia de emprego durante o per�odo que faltar at� a aquisi��o do direito<font style="text-decoration:line-through;font-weight:bold">, exceto nos cargos de confian�a ou de mandato com dura��o expressa de inicio e t�rmino</font>.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	24. Garantias ao Auxiliar em vias de Aposentadoria
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo primeiro � A garantia de emprego � devida ao AUXILIAR que esteja contratado pela MANTENEDORA h� pelo menos tr�s anos. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo primeiro - A garantia de emprego � devida ao AUXILIAR que esteja contratado pela MANTENEDORA h� pelo menos tr�s anos<font style="text-decoration:line-through;font-weight:bold"> e que tenha comunicado � mesma a solicita��o de sua contagem de tempo</font>.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	28. Indeniza��o por Dispensa Imotivada
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo primeiro � N�o ter� direito a indeniza��o prevista na al�nea �a� o AUXILIAR que tiver recebido, durante pelo menos um ano, pagamento mensal de adicional por tempo de servi�o decorrente de plano de cargos e sal�rios ou de anu�nio, q�inq��nio ou equivalente, cujo valor corresponda a, no m�nimo, 1% (um por cento) do valor do sal�rio, por ano trabalhado. <b>A MANTENEDORA dever� apresentar, no momento da homologa��o, documentos que comprovem o pagamento ao AUXILIAR do referido adicional por tempo de servi�o. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo primeiro - N�o estar� obrigada ao pagamento da indeniza��o, prevista na al�nea �a�, a MANTENEDORA que tiver garantido ao AUXILIAR demitido, durante pelo menos um ano, pagamento mensal de adicional por tempo de servi�o decorrente de plano de cargos e sal�rios ou de anu�nio, q�inq��nio ou equivalente, cujo valor corresponda a, no m�nimo, 1% do valor do sal�rio por ano trabalhado.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	48. Cesta B�sica
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada aos AUXILIARES que percebam, at� 5 (cinco) sal�rios m�nimos por m�s, em jornada de <b>36 (trinta e seis) horas semanais, ou percebam, em jornada inferior, remunera��o proporcionalmente igual ou inferior ao limite fixado</b> nesta cl�usula, a concess�o de uma cesta b�sica mensal de 26 kg, composta, no m�nimo, dos seguintes produtos n�o perec�veis: 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada aos AUXILIARES que percebam, at� 5 (cinco) sal�rios m�nimos por m�s, em jornada integral de <b>44 (quarenta e quatro) horas semanais</b>, a concess�o de uma cesta b�sica mensal de 26 kg, composta, no m�nimo, dos seguintes produtos n�o perec�veis:
	</td></tr>
</table>

<DIV style="page-break-after:always"></DIV>

<!-- professores -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td valign=top class=grupo colspan=3>Professores</td></tr>
<tr><td valign=top class=titulo>Cl�usula</td>
	<td valign=top class=titulo>Teor atual</td>
	<td valign=top class=titulo>Teor anterior</td></tr>

<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	1. Abrang�ncia
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo segundo � Quando o PROFESSOR for contratado em um munic�pio para exercer a sua atividade em outro, prevalecer� o cumprimento da Conven��o Coletiva do munic�pio onde o servi�o � prestado. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	18. Atestados M�dicos e Abono de Faltas
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	A MANTENEDORA ser� obrigada a abonar as faltas dos PROFESSORES, mediante a apresenta��o de atestados m�dicos ou odontol�gicos. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	A MANTENEDORA est� obrigada a aceitar atestados fornecidos por m�dicos ou dentistas <font style="text-decoration:line-through;font-weight:bold">credenciados pelo SINPRO, SUS ou, ainda, profissionais conveniados com a pr�pria MANTENEDORA.
	<br>Par�grafo �nico � Tamb�m ser�o aceitos atestados que tenham sido convalidados pelos profissionais de sa�de do departamento m�dico ou odontol�gico do SINPRO ou conveniados a ele.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	23. Abono de Faltas por Casamento ou Luto
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo �nico � N�o ser�o descontadas, no curso de 3 (tr�s) dias, as faltas do PROFESSOR por motivo de falecimento de sogra, sogro, neto, neta, irm� ou irm�o. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	32. Garantias ao Professor em Vias de Aposentadoria
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo segundo � A comprova��o � MANTENEDORA dever� ser feita mediante a apresenta��o de documento que ateste o tempo de servi�o. Este documento dever� ser emitido por pessoa credenciada junto ao �rg�o previdenci�rio. Se o PROFESSOR depender de documenta��o para realiza��o da contagem, <b>ter� um prazo de 30 (trinta) dias, a contar da data prevista ou marcada para homologa��o da rescis�o contratual</b>. Comprovada a solicita��o de tal documenta��o, os prazos ser�o prorrogados at� que a mesma seja emitida, <b>assegurando-se, nessa situa��o, o pagamento dos sal�rios pelo prazo m�ximo de 120 dias</b>. 
	<br>
	Par�grafo sexto � Para garantir a estabilidade prevista nesta cl�usula, o professor dever� encaminhar � MANTENEDORA, dentro da prorroga��o prevista no par�grafo 2�, documenta��o que demonstre a tramita��o do processo que atesta o tempo de servi�o. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo segundo � A comprova��o � MANTENEDORA dever� ser feita mediante a apresenta��o de documento que ateste o tempo de servi�o. Este documento dever� ser emitido pela Previd�ncia Social ou por pessoa credenciada junto ao �rg�o previdenci�rio. Se o PROFESSOR depender de documenta��o para realiza��o da contagem, <b>ter� um prazo de quarenta e cinco dias, a contar da data da comunica��o da dispensa</b>. Comprovada a solicita��o de tal documenta��o, os prazos ser�o prorrogados at� que a mesma seja emitida.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	36. Indeniza��es por Dispensa Imotivada
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo primeiro � N�o ter� direito � indeniza��o assegurada na al�nea a) do caput o PROFESSOR que tiver recebido, durante pelo menos um ano, pagamento mensal de adicional por tempo de servi�o decorrente de plano de cargos e sal�rios ou de anu�nio, q�inq��nio ou equivalente, cujo valor corresponda a, no m�nimo, 1% (um por cento) do valor da hora-aula por ano trabalhado e, por conseq��ncia, do sal�rio mensal. <b>A MANTENEDORA dever� apresentar, no momento da homologa��o, documentos que comprovem o pagamento ao PROFESSOR do referido adicional por tempo de servi�o</b>. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo primeiro � N�o estar� obrigada ao pagamento da indeniza��o prevista na al�nea a) a MANTENEDORA> que tiver garantido ao PROFESSOR demitido, durante pelo menos um ano, pagamento mensal de adicional por tempo de servi�o decorrente de plano de cargos e sal�rios ou de anu�nio, q�inq��nio ou equivalente, cujo valor corresponda a, no m�nimo, 1% (um por cento) do valor da hora-aula por ano trabalhado e, por conseq��ncia, do sal�rio mensal.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	39. Recesso Escolar
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo quarto � Os calend�rios escolares que definir�o os per�odos de recesso escolar dos PROFESSORES ser�o obrigatoriamente divulgados aos PROFESSORES at� o in�cio de cada per�odo letivo<b> e enviados ao SINPRO</b>. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo quarto � Os calend�rios escolares que definir�o os per�odos de recesso escolar dos PROFESSORES ser�o obrigatoriamente divulgados aos PROFESSORES at� o in�cio de cada per�odo letivo.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	41. Quadro de Avisos
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo �nico � O dirigente sindical ter� livre acesso � sala dos PROFESSORES, no hor�rio de intervalo das aulas, para atualiza��o do material divulgado no quadro de avisos, uma �nica vez em cada m�s. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	57. Disposi��es Transit�rias
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica estabelecido que a FEPESP, os Sindicatos de Professores integrantes, o SEMESP e o SEMESP�RIO PRETO constituir�o uma comiss�o, denominada �Comiss�o de Aprimoramento das Rela��es de Trabalho�, composta, de forma parit�ria, por 4 representantes de cada uma das categorias, profissional e econ�mica, que dever� reunir-se, ordin�ria e obrigatoriamente, mensalmente, entre maio e outubro de 2007 e, extraordinariamente, sempre que convocada por, no m�nimo, 5 (cinco) de seus membros, com a pauta espec�fica de discutir os seguintes temas de interesse de ambas as categorias: 
	<br>a) rela��es de trabalho envolvendo aplica��es de novas tecnologias, ensino � dist�ncia, cursos semi-presenciais e tele-presenciais; 
	<br>b) rela��es de trabalho nos cursos modulares e seq�enciais; 
	<br>c) planos de carreira das Institui��es privadas de ensino; 
	<br>d) atividade docente, pesquisadores, orientadores, coordenadores de �reas, disciplinas, departamentos, etc. 
	<br>e) Assist�ncia M�dico-Hospitalar, no que se refere � sua eventual implementa��o por interm�dio das entidades sindicais profissionais. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
</table>

<%
'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>
<!-- -->
<!-- -->
</body>
</html>