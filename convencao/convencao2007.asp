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
<title>Pontos importantes da Convenção Coletiva 2007</title>
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
<tr><td valign=top class=titulo colspan=3>Pontos importantes das Convenções Coletivas 2007</td></tr>
<tr><td valign=top class=grupo colspan=3>Auxiliares</td></tr>
<tr><td valign=top class=titulo>Cláusula</td>
	<td valign=top class=titulo>Teor atual</td>
	<td valign=top class=titulo>Teor anterior</td></tr>

<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	13. Anotações na Carteira de Trabalho
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo único – É obrigatória a anotação na CTPS das mudanças provocadas por ascensão em plano de carreira. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo único - É obrigatória a anotação na CTPS das mudanças provocadas por ascensão em plano de carreira <font style="text-decoration:line-through;font-weight:bold">ou alteração de titulação</font>.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	16. Bolsas de Estudo
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Todo AUXILIAR tem direito a bolsas de estudo integrais, incluindo matrícula, para si, <b>cônjuge</b>, filhos ou dependentes legais, ambos entendidos como aqueles reconhecidos pela legislação do Imposto de Renda ou aqueles que estejam sob a guarda judicial do AUXILIAR e vivam sob sua dependência econômica, devidamente comprovada.	
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo 1º - Somente terão direito a bolsas de estudo integrais, o(a) AUXILIAR, <font style="text-decoration:line-through;font-weight:bold">esposo(a) e companheiro(a)</font>, bem como seus filhos(as) e dependentes legais que estejam sob a guarda judicial, estes dois últimos desde que tenham 25 (vinte e cinco) anos ou menos na data de realização do exame vestibular ou do processo seletivo que define o ingresso no curso superior.	
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	16. Bolsas de Estudo
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo sétimo – As bolsas de estudo integrais em cursos de pós-graduação ou especialização existentes e administrados pela MANTENEDORA são válidas exclusivamente para o AUXILIAR, respeitados os critérios de seleção exigidos para ingresso nos mesmos e obedecerão às seguintes condições: 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo oitavo - As bolsas de estudo integrais em cursos de pós-graduação ou de especialização existentes e administrados pela MANTENEDORA são válidas exclusivamente para o AUXILIAR <font style="text-decoration:line-through;font-weight:bold">em áreas correlatas àquelas em que o AUXILIAR exerce a função na MANTENEDORA e que visem à sua capacitação,</font> respeitados os critérios de seleção exigidos para ingresso nos mesmos e obedecerão às seguintes condições:
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	16. Bolsas de Estudo
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo onze - Quando, a critério da MANTENEDORA, o AUXILIAR, em razão das funções exercidas na Instituição se vir na contingência de efetuar seus estudos, na área educacional indicada em outra instituição de ensino, a MANTENEDORA arcará com o valor integral das mensalidades do curso, incluindo matrícula durante a vigência do contrato de trabalho, respeitada a vigência coletiva de trabalho.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	23. Creches	
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	É obrigatória a instalação de local destinado à guarda de crianças de até <b>doze meses</b>, quando a unidade de ensino da MANTENEDORA mantiver contratadas, em jornada integral, pelo menos trinta funcionárias. A manutenção da creche poderá ser substituída pelo pagamento do reembolso-creche, nos termos da legislação em vigor (CF, 7º, XXV, Artigo 389, parágrafo 1º da CLT e Portaria MTb nº 3296 de 03.09.86), ou ainda, a celebração de convênio com uma entidade reconhecidamente idônea. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	É obrigatória a instalação de local destinado à guarda de crianças de até <font style="text-decoration:line-through;font-weight:bold">seis anos</font>, quando a unidade de ensino da mantiver contratadas, em jornada integral, pelo menos trinta funcionárias<font style="text-decoration:line-through;font-weight:bold"> com idade superior a 16 anos</font>. A manutenção da creche poderá ser substituída pelo pagamento do reembolso-creche, nos termos da legislação em vigor (CF, 7º, XXV, Artigo 389, parágrafo 1º da CLT e Portaria MTb nº 3296 de 03.09.86), ou ainda, a celebração de convênio com uma entidade reconhecidamente idônea.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	24. Garantias ao Auxiliar em vias de Aposentadoria
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurado ao AUXILIAR que, comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria por tempo de contribuição ou da aposentadoria por idade, a garantia de emprego durante o período que faltar até a aquisição do direito. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada ao AUXILIAR que, comprovadamente estiver a 24 meses ou menos da aposentadoria integral por tempo de serviço ou da aposentadoria por idade, a garantia de emprego durante o período que faltar até a aquisição do direito<font style="text-decoration:line-through;font-weight:bold">, exceto nos cargos de confiança ou de mandato com duração expressa de inicio e término</font>.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	24. Garantias ao Auxiliar em vias de Aposentadoria
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo primeiro – A garantia de emprego é devida ao AUXILIAR que esteja contratado pela MANTENEDORA há pelo menos três anos. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo primeiro - A garantia de emprego é devida ao AUXILIAR que esteja contratado pela MANTENEDORA há pelo menos três anos<font style="text-decoration:line-through;font-weight:bold"> e que tenha comunicado à mesma a solicitação de sua contagem de tempo</font>.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	28. Indenização por Dispensa Imotivada
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo primeiro – Não terá direito a indenização prevista na alínea “a” o AUXILIAR que tiver recebido, durante pelo menos um ano, pagamento mensal de adicional por tempo de serviço decorrente de plano de cargos e salários ou de anuênio, qüinqüênio ou equivalente, cujo valor corresponda a, no mínimo, 1% (um por cento) do valor do salário, por ano trabalhado. <b>A MANTENEDORA deverá apresentar, no momento da homologação, documentos que comprovem o pagamento ao AUXILIAR do referido adicional por tempo de serviço. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo primeiro - Não estará obrigada ao pagamento da indenização, prevista na alínea “a”, a MANTENEDORA que tiver garantido ao AUXILIAR demitido, durante pelo menos um ano, pagamento mensal de adicional por tempo de serviço decorrente de plano de cargos e salários ou de anuênio, qüinqüênio ou equivalente, cujo valor corresponda a, no mínimo, 1% do valor do salário por ano trabalhado.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	48. Cesta Básica
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada aos AUXILIARES que percebam, até 5 (cinco) salários mínimos por mês, em jornada de <b>36 (trinta e seis) horas semanais, ou percebam, em jornada inferior, remuneração proporcionalmente igual ou inferior ao limite fixado</b> nesta cláusula, a concessão de uma cesta básica mensal de 26 kg, composta, no mínimo, dos seguintes produtos não perecíveis: 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada aos AUXILIARES que percebam, até 5 (cinco) salários mínimos por mês, em jornada integral de <b>44 (quarenta e quatro) horas semanais</b>, a concessão de uma cesta básica mensal de 26 kg, composta, no mínimo, dos seguintes produtos não perecíveis:
	</td></tr>
</table>

<DIV style="page-break-after:always"></DIV>

<!-- professores -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td valign=top class=grupo colspan=3>Professores</td></tr>
<tr><td valign=top class=titulo>Cláusula</td>
	<td valign=top class=titulo>Teor atual</td>
	<td valign=top class=titulo>Teor anterior</td></tr>

<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	1. Abrangência
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo segundo – Quando o PROFESSOR for contratado em um município para exercer a sua atividade em outro, prevalecerá o cumprimento da Convenção Coletiva do município onde o serviço é prestado. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	18. Atestados Médicos e Abono de Faltas
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	A MANTENEDORA será obrigada a abonar as faltas dos PROFESSORES, mediante a apresentação de atestados médicos ou odontológicos. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	A MANTENEDORA está obrigada a aceitar atestados fornecidos por médicos ou dentistas <font style="text-decoration:line-through;font-weight:bold">credenciados pelo SINPRO, SUS ou, ainda, profissionais conveniados com a própria MANTENEDORA.
	<br>Parágrafo único – Também serão aceitos atestados que tenham sido convalidados pelos profissionais de saúde do departamento médico ou odontológico do SINPRO ou conveniados a ele.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	23. Abono de Faltas por Casamento ou Luto
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo único – Não serão descontadas, no curso de 3 (três) dias, as faltas do PROFESSOR por motivo de falecimento de sogra, sogro, neto, neta, irmã ou irmão. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	32. Garantias ao Professor em Vias de Aposentadoria
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo segundo – A comprovação à MANTENEDORA deverá ser feita mediante a apresentação de documento que ateste o tempo de serviço. Este documento deverá ser emitido por pessoa credenciada junto ao órgão previdenciário. Se o PROFESSOR depender de documentação para realização da contagem, <b>terá um prazo de 30 (trinta) dias, a contar da data prevista ou marcada para homologação da rescisão contratual</b>. Comprovada a solicitação de tal documentação, os prazos serão prorrogados até que a mesma seja emitida, <b>assegurando-se, nessa situação, o pagamento dos salários pelo prazo máximo de 120 dias</b>. 
	<br>
	Parágrafo sexto – Para garantir a estabilidade prevista nesta cláusula, o professor deverá encaminhar à MANTENEDORA, dentro da prorrogação prevista no parágrafo 2º, documentação que demonstre a tramitação do processo que atesta o tempo de serviço. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo segundo – A comprovação à MANTENEDORA deverá ser feita mediante a apresentação de documento que ateste o tempo de serviço. Este documento deverá ser emitido pela Previdência Social ou por pessoa credenciada junto ao órgão previdenciário. Se o PROFESSOR depender de documentação para realização da contagem, <b>terá um prazo de quarenta e cinco dias, a contar da data da comunicação da dispensa</b>. Comprovada a solicitação de tal documentação, os prazos serão prorrogados até que a mesma seja emitida.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	36. Indenizações por Dispensa Imotivada
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo primeiro – Não terá direito à indenização assegurada na alínea a) do caput o PROFESSOR que tiver recebido, durante pelo menos um ano, pagamento mensal de adicional por tempo de serviço decorrente de plano de cargos e salários ou de anuênio, qüinqüênio ou equivalente, cujo valor corresponda a, no mínimo, 1% (um por cento) do valor da hora-aula por ano trabalhado e, por conseqüência, do salário mensal. <b>A MANTENEDORA deverá apresentar, no momento da homologação, documentos que comprovem o pagamento ao PROFESSOR do referido adicional por tempo de serviço</b>. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo primeiro – Não estará obrigada ao pagamento da indenização prevista na alínea a) a MANTENEDORA> que tiver garantido ao PROFESSOR demitido, durante pelo menos um ano, pagamento mensal de adicional por tempo de serviço decorrente de plano de cargos e salários ou de anuênio, qüinqüênio ou equivalente, cujo valor corresponda a, no mínimo, 1% (um por cento) do valor da hora-aula por ano trabalhado e, por conseqüência, do salário mensal.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	39. Recesso Escolar
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo quarto – Os calendários escolares que definirão os períodos de recesso escolar dos PROFESSORES serão obrigatoriamente divulgados aos PROFESSORES até o início de cada período letivo<b> e enviados ao SINPRO</b>. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo quarto – Os calendários escolares que definirão os períodos de recesso escolar dos PROFESSORES serão obrigatoriamente divulgados aos PROFESSORES até o início de cada período letivo.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	41. Quadro de Avisos
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo único – O dirigente sindical terá livre acesso à sala dos PROFESSORES, no horário de intervalo das aulas, para atualização do material divulgado no quadro de avisos, uma única vez em cada mês. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	57. Disposições Transitórias
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica estabelecido que a FEPESP, os Sindicatos de Professores integrantes, o SEMESP e o SEMESP–RIO PRETO constituirão uma comissão, denominada “Comissão de Aprimoramento das Relações de Trabalho”, composta, de forma paritária, por 4 representantes de cada uma das categorias, profissional e econômica, que deverá reunir-se, ordinária e obrigatoriamente, mensalmente, entre maio e outubro de 2007 e, extraordinariamente, sempre que convocada por, no mínimo, 5 (cinco) de seus membros, com a pauta específica de discutir os seguintes temas de interesse de ambas as categorias: 
	<br>a) relações de trabalho envolvendo aplicações de novas tecnologias, ensino à distância, cursos semi-presenciais e tele-presenciais; 
	<br>b) relações de trabalho nos cursos modulares e seqüenciais; 
	<br>c) planos de carreira das Instituições privadas de ensino; 
	<br>d) atividade docente, pesquisadores, orientadores, coordenadores de áreas, disciplinas, departamentos, etc. 
	<br>e) Assistência Médico-Hospitalar, no que se refere à sua eventual implementação por intermédio das entidades sindicais profissionais. 
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