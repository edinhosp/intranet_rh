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
<title>Conven��o Coletiva 2013/14 - Professores</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<!-- -->
<table><tr><td>
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
<tr><td class=titulo align="center">CONVEN��O COLETIVA DE TRABALHO PARA 2013/14
<tr><td class=titulo align="center">SEMESP
<tr><td class=titulo align="center">PROFESSORES 
<tr><td class=campo style="text-align:justify">

<tr><td class=titulo>1. Abrang�ncia
<tr><td class=campo style="text-align:justify">Esta Conven��o abrange a categoria econ�mica dos estabelecimentos particulares de ensino superior no Estado de S�o Paulo, aqui designados como MANTENEDORA e a categoria profissional diferenciada dos professores, aqui designada simplesmente como PROFESSOR.
<br><b>Par�grafo �nico</b> � A categoria dos PROFESSORES abrange todos aqueles que exercem a atividade docente, independentemente da denomina��o sob a qual a fun��o for exercida. Considera-se atividade docente a fun��o de ministrar aulas.

<tr><td class=titulo>2. Dura��o
<tr><td class=campo style="text-align:justify">Esta Conven��o Coletiva de Trabalho ter� dura��o de dois anos, com vig�ncia de 1� de mar�o de 2013 a 28 de fevereiro de 2015.
<br><b>Par�grafo �nico</b> � As cl�usulas poder�o ser reexaminadas na pr�xima data base, em 1� de mar�o de 2014, em virtude de problemas surgidos na sua aplica��o ou do surgimento de normas legais a elas pertinentes, ou em decorr�ncia de aprova��o das propostas apresentadas pela Comiss�o de Aprimoramento das Rela��es de Trabalho prevista na presente Conven��o.

<tr><td class=titulo>3. Reajuste salarial em 1� de mar�o de 2013
<tr><td class=campo style="text-align:justify">A partir de 1� de mar�o de 2013, ser� aplicado o reajuste de 6,52%, sobre os sal�rios devidos em 1� de agosto de 2012.
<br><b>Par�grafo �nico</b> � Fica estabelecido que o sal�rio de 1� de mar�o de 2013, reajustado pelo �ndice definido nesta cl�usula, servir� como base de c�lculo para a data base de 1� de mar�o de 2014.

<tr><td class=titulo>4. Reajuste salarial em 1� de mar�o de 2014
<tr><td class=campo style="text-align:justify">Em 1� de mar�o de 2014, as MANTENEDORAS dever�o aplicar sobre os sal�rios devidos em 1� de mar�o de 2013, o percentual definido pela m�dia aritm�tica dos �ndices inflacion�rios do per�odo compreendido entre 1� de mar�o de 2013 e 28 de fevereiro de 2014, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV), acrescido de 1%, a t�tulo de aumento real.
<br><b>Par�grafo primeiro</b> � Fica estabelecido que o sal�rio de 1� de mar�o de 2014, reajustado pelo �ndice definido nesta cl�usula, servir� como base de c�lculo para a data base de 1� de mar�o de 2015.
<br><b>Par�grafo segundo</b> � O SEMESP, o SINDICATO e a FEPESP comprometem-se a divulgar, em comunicado conjunto, at� 20 de mar�o de 2014, o percentual de reajuste salarial calculado pela f�rmula definida no caput.

<tr><td class=titulo>5. Compensa��es salariais
<tr><td class=campo style="text-align:justify">No ano de 2013 ser� permitida a compensa��o de eventuais antecipa��es salariais concedidas no per�odo compreendido entre 1� de mar�o de 2012 e 28 de fevereiro de 2013. Relativamente � database de mar�o de 2014 ser� permitida a compensa��o de eventuais antecipa��es salariais concedidas no per�odo compreendido entre 1� de mar�o de 2013 e 28 de fevereiro de 2014.
<br><b>Par�grafo �nico</b> � N�o ser� permitida, em ambos os casos, a compensa��o daquelas antecipa��es salariais que decorrerem de promo��es, transfer�ncias, ascens�o em plano de carreira e os reajustes concedidos com cl�usula expressa de n�o compensa��o.

<tr><td class=titulo>6. Composi��o do sal�rio mensal do professor
<tr><td class=campo style="text-align:justify">O sal�rio do PROFESSOR � composto, no m�nimo, por tr�s itens: o sal�rio base, o descanso semanal remunerado (DSR) e a hora-atividade.
<br>O sal�rio base � calculado pela seguinte equa��o: n�mero de aulas semanais multiplicado por 4,5 semanas e multiplicado, ainda, pelo valor da hora-aula (artigo 320, Par�grafo 1� da CLT).
<br>O DSR corresponde a 1/6 do sal�rio base, acrescido, quando houver, do total de horas extras e do adicional noturno (Lei 605/49).
<br>A hora-atividade corresponde a 5% do total obtido com a somat�ria de todos os valores acima referidos.
<br><b>Par�grafo �nico</b> - A remunera��o adicional do PROFESSOR pelo exerc�cio concomitante de fun��o n�o-docente obedecer� aos crit�rios estabelecidos entre a MANTENEDORA e o PROFESSOR que aceitar o cargo.

<tr><td class=titulo>7. Prazo para pagamento de sal�rios
<tr><td class=campo style="text-align:justify">Os sal�rios dever�o ser pagos, no m�ximo, at� o quinto dia �til do m�s subsequente ao trabalhado.
<br><b>Par�grafo �nico</b> - O n�o pagamento dos sal�rios no prazo obriga a MANTENEDORA a pagar multa di�ria, em favor do PROFESSOR, no valor de 1/50 de seu sal�rio mensal.

<tr><td class=titulo>8. Comprovante de pagamento
<tr><td class=campo style="text-align:justify">A MANTENEDORA dever� fornecer ao PROFESSOR, mensalmente, comprovante de pagamento, devendo estar discriminados: 
<blockquote style="margin-top:0;margin-bottom:0">
	a) identifica��o da MANTENEDORA e do estabelecimento de ensino;
	<br>b) a identifica��o do PROFESSOR; 
	<br>c) a denomina��o da categoria e, se houver, faixas salariais diferenciadas, inclusive aquelas definidas em eventual plano de carreira da Institui��o; 
	<br>d) o valor da hora-aula; 
	<br>e) a carga hor�ria semanal; 
	<br>f) a hora-atividade; 
	<br>g) outros eventuais adicionais, inclusive o adicional por tempo de servi�o, caso exista; 
	<br>h) o descanso semanal remunerado; 
	<br>i) as horas extras realizadas; 
	<br>j) o valor do recolhimento do FGTS; 
	<br>l) o desconto previdenci�rio; 
	<br>m) outros descontos.
</blockquote>

<tr><td class=titulo>9. Autoriza��o para desconto em folha de pagamento
<tr><td class=campo style="text-align:justify">O desconto do professor em folha de pagamento somente poder� ser realizado mediante sua autoriza��o, nos termos dos artigos 462 e 545 da CLT, quando os valores forem destinados ao custeio de pr�mios de seguro, planos de sa�de, mensalidades associativas ou outras que constem da sua expressa autoriza��o, desde que n�o haja previs�o expressa de desconto na presente norma coletiva.
<br><b>Par�grafo �nico</b> � Encontra-se no Sindicato, � disposi��o da MANTENEDORA, devendo ser a ela encaminhada, quando solicitada formalmente, c�pia de autoriza��o do PROFESSOR para o desconto da mensalidade associativa.

<tr><td class=titulo>10. Horas extras
<tr><td class=campo style="text-align:justify">Considera-se atividade extra todo trabalho desenvolvido em hor�rio diferente daquele habitualmente realizado na semana. As atividades extras devem ser pagas com adicional de 100%.
<br><b>Par�grafo primeiro</b> � N�o � considerada atividade extra a participa��o em cursos de capacita��o e aperfei�oamento docente, desde que aceita livremente pelo PROFESSOR.
<br><b>Par�grafo segundo</b> � Ser�o pagas apenas como aulas normais, acrescidas do DSR e da horaatividade, aquelas que forem adicionadas provisoriamente � carga hor�ria habitual, decorrentes:
<blockquote style="margin-top:0;margin-bottom:0">
	a) da substitui��o tempor�ria de outro PROFESSOR, com dura��o predeterminada, decorrente de licen�a m�dica, maternidade ou para estudos. Nestes casos, a substitui��o dever� ser formalizada atrav�s de documento firmado entre a MANTENEDORA e o PROFESSOR que aceitar realiz�-la;
	<br>b) de substitui��es eventuais de faltas de PROFESSOR respons�vel, desde que aceitas livremente pelo PROFESSOR substituto;
	<br>c) de reposi��o de eventuais faltas que foram descontadas dos sal�rios nos meses em que ocorreram;
	<br>d) da realiza��o de cursos eventuais ou de curta dura��o, inclusive cursos de depend�ncia, e aceitas livremente, mediante documento firmado entre o PROFESSOR convidado a ministr�-los e a MANTENEDORA.
	<br>e) do comparecimento a reuni�es did�tico-pedag�gicas, de avalia��o e de planejamento, quando realizadas fora de seu hor�rio habitual de trabalho, desde que aceito livremente pelo PROFESSOR.
</blockquote>
<br><b>Par�grafo terceiro</b> � A participa��o em Comiss�es Internas e Externas da Unidade de Ensino da MANTENEDORA, desde que aceita livremente pelo PROFESSOR mediante documento firmado, ser� remunerada como aula ou hora normal, acrescida de DSR.

<tr><td class=titulo>11. Adicional noturno
<tr><td class=campo style="text-align:justify">O trabalho noturno deve ser pago nas atividades realizadas ap�s as 22 horas e corresponde a 25% do valor da hora-aula.

<tr><td class=titulo>12. Hora-atividade
<tr><td class=campo style="text-align:justify">Fica mantido o adicional de 5% a t�tulo de hora-atividade, destinado exclusivamente ao pagamento do tempo gasto pelo PROFESSOR, fora do estabelecimento de ensino, na prepara��o de aulas, provas e exerc�cios, bem como na corre��o dos mesmos.

<tr><td class=titulo>13. Adicional por atividades em outros munic�pios
<tr><td class=campo style="text-align:justify">Quando o PROFESSOR desenvolver suas atividades a servi�o da mesma MANTENEDORA em munic�pio diferente daquele onde foi contratado e onde ocorre a presta��o habitual do trabalho, dever� receber um adicional de 25% sobre o total de sua remunera��o no novo munic�pio. Quando o PROFESSOR voltar a prestar servi�os no munic�pio de origem, cessar� a obriga��o no pagamento do adicional.
<br><b>Par�grafo primeiro</b> - Nos casos em que ocorrer a transfer�ncia definitiva do PROFESSOR, aceita livremente por este, em documento firmado entre as partes, n�o haver� a incid�ncia do adicional referido no caput, obrigando-se a MANTENEDORA a efetuar o pagamento de um �nico sal�rio mensal integral, ao PROFESSOR, no ato da transfer�ncia, a t�tulo de ajuda de custo.
<br><b>Par�grafo segundo</b> - Fica assegurada a garantia de emprego pelo per�odo de seis meses ao PROFESSOR transferido de munic�pio, contados a partir do in�cio do trabalho e/ou da efetiva��o da transfer�ncia.
<br><b>Par�grafo terceiro</b> � Caso a MANTENEDORA desenvolva atividade acad�mica em munic�pios considerados conurbados, poder� solicitar isen��o do pagamento do adicional determinado no caput, desde que encaminhe material comprobat�rio ao SEMESP, para an�lise e delibera��o do Foro Conciliat�rio para Solu��o de Conflitos Coletivos, previsto na presente Conven��o.

<tr><td class=titulo>14. Participa��o nos lucros ou resultados ou abono especial
<tr><td class=campo style="text-align:justify">Na vig�ncia da presente Conven��o Coletiva de Trabalho, obriga-se a MANTENEDORA a pagar aos seus PROFESSORES a parcela correspondente a 24% do seu sal�rio mensal bruto, at� o dia 15 de outubro de 2014. Tal pagamento poder� ser feito a t�tulo de participa��o nos lucros ou resultados (PLR) ou como abono especial, segundo os crit�rios e normas especificadas nos Par�grafos seguintes:
<br><b>Par�grafo primeiro</b> � A MANTENEDORA pagar� a seus PROFESSORES a parcela correspondente a 24% do seu sal�rio mensal bruto, a t�tulo de participa��o nos lucros ou resultados (PLR), nos termos da Lei n� 10.101, de 19 de dezembro de 2000, com as altera��es da Lei n� 12.832, de 20 de junho de 2013, caso a Institui��o de Ensino Superior mantida tenha alcan�ado pelo menos dois dos requisitos abaixo nominados:
<blockquote style="margin-top:0;margin-bottom:0">
	a) Redu��o do �ndice de inadimpl�ncia no ano de 2013, em rela��o ao ano anterior;
	<br>b) Pelo menos 70% positivos nos indicadores de avalia��o publicados pelo Minist�rio da Educa��o;
	<br>c) Redu��o da evas�o no ano de 2013, em rela��o ao ano anterior;
	<br>d) �ndice satisfat�rio na maioria dos indicadores dos relat�rios de auto-avalia��o produzidos pela CPA.
</blockquote>
<br><b>Par�grafo segundo</b> � Os requisitos acima definidos, assim como o percentual do sal�rio pago aos PROFESSORES a t�tulo de participa��o nos lucros ou resultados (PLR) poder�o ser substitu�dos pela MANTENEDORA, desde que seja garantido o m�nimo de 24% do sal�rio bruto e a obten��o de no m�ximo 50% das novas metas definidas e que o SEMESP e a FEPESP tenham sido comunicados da altera��o, at� 30 de junho de 2014, sujeitos � aprova��o do Foro Conciliat�rio de Solu��o de Conflitos Coletivos, em reuni�o especialmente convocada para esse prop�sito.
<br><b>Par�grafo terceiro</b> � A MANTENEDORA que n�o conceder a participa��o nos lucros ou resultados (PLR) estabelecida no caput, seja em raz�o do n�o atendimento dos requisitos e condi��es definidas nesta cl�usula, ou em raz�o de se considerar enquadrada no inciso II do Par�grafo 3� do artigo 2� da Lei n.� 10.101/2000 com as altera��es da Lei n� 12.832, de 20 de junho de 2013, ou ainda, em raz�o de outro motivo qualquer, dever� pagar a seus PROFESSORES, no prazo acima definido, a parcela correspondente a 24% do seu sal�rio mensal bruto, a t�tulo de abono salarial.

<tr><td class=titulo>15. Bolsas de estudo
<tr><td class=campo style="text-align:justify">Todo PROFESSOR tem direito a bolsa de estudo integral, incluindo matr�cula, em cursos de gradua��o, sequenciais e p�s-gradua��o existentes e administrados pela MANTENEDORA que o emprega, observado o que segue:
<blockquote style="margin-top:0;margin-bottom:0">
	1. A MANTENEDORA est� obrigada a conceder, no m�ximo, duas bolsas de estudo, sendo que, nos cursos de gradua��o e sequenciais, n�o ser� poss�vel que o PROFESSOR conclua mais de um curso nessa condi��o.
	<br>2. As bolsas de estudo integrais em cursos de p�s-gradua��o ou especializa��o existentes e administrados pela MANTENEDORA s�o v�lidas exclusivamente para o PROFESSOR, em �reas correlatas �s disciplinas que o mesmo ministra na Institui��o e que visem a capacita��o docente, respeitados os crit�rios de sele��o exigidos para ingresso no mesmo e obedecer�o as seguintes condi��es :
<blockquote style="margin-top:0;margin-bottom:0">
		a) nos cursos stricto sensu ou de especializa��o que fixem um n�mero m�ximo de alunos por turma, s�o limitadas em 30% do total de vagas oferecidas;
		<br>b) nos cursos de p�s-gradua��o lato sensu n�o haver� limites de vagas. Caso a estrutura do curso torne necess�ria a limita��o do n�mero de alunos ser� observado o disposto na al�nea �a� deste item.
</blockquote>
	3. O direito �s bolsas de estudo passa a vigorar ao t�rmino do contrato de experi�ncia, cuja dura��o n�o pode exceder de noventa dias, conforme Par�grafo �nico do artigo 445 da CLT.
	<br>4. As bolsas de estudo ser�o mantidas quando o PROFESSOR estiver licenciado para tratamento de sa�de ou em gozo de licen�a mediante anu�ncia da MANTENEDORA, excetuado o disposto na cl�usula �Licen�a sem Remunera��o�.
	<br>5. O PROFESSOR que for reprovado no per�odo letivo perder� o direito � bolsa de estudo, voltando a gozar do benef�cio quando lograr aprova��o no referido per�odo. As disciplinas cursadas em regime de depend�ncia ser�o de total responsabilidade do PROFESSOR, arcando o mesmo com o seu custo.
</blockquote>

<tr><td class=campo style="text-align:justify">B - Programa de Inclus�o, Capacita��o para Filhos, Dependentes Legais e Estudantes 
<br>O CEBRADE � Centro Brasileiro de Desenvolvimento do Ensino Superior � tem, como um dos seus objetivos, desenvolver o Programa de Amparo Educativo Tempor�rio � PAET, concedendo bolsas de estudo em Institui��es Privadas de Ensino Superior. Os filhos ou dependentes legais do PROFESSOR t�m direito a usufruir as gratuidades integrais do PAET, sem qualquer �nus, nos cursos de gradua��o ou sequenciais existentes e administrados pela MANTENEDORA para a qual o PROFESSOR trabalha, observado o disposto nesta cl�usula e no �Regulamento do Programa de Capacita��o�, anexado � presente Conven��o.
<br><b>Par�grafo primeiro</b> � A MANTENEDORA dever� disponibilizar ao CEBRADE, mediante requerimento, bolsas de estudo em n�mero suficiente para o atendimento da concess�o das gratuidades integrais do PAET nas Institui��es de Ensino Superior por ela mantida, para filhos ou dependentes legais dos seus PROFESSORES, observada a limita��o de duas bolsas de estudo por PROFESSOR.
<br><b>Par�grafo segundo</b> � O benefici�rio bolsista, concluinte de curso de gradua��o ou seq�encial, n�o poder� obter nova concess�o de gratuidade em um desses cursos, na mesma IES.
<br><b>Par�grafo terceiro</b> � O SEMESP e a FEDERA��O representante da categoria profissional fiscalizar�o o CEBRADE na gest�o do Programa de Amparo Educativo Tempor�rio para os filhos e dependentes legais dos PROFESSORES, na conformidade do estabelecido nesta cl�usula e no �Regulamento do Programa de Capacita��o�.
<br><b>Par�grafo quarto</b> � Para a concess�o das gratuidades integrais aos filhos e dependentes legais do PROFESSOR, o CEBRADE n�o poder� fazer qualquer outra exig�ncia a n�o ser o comprovante de aprova��o no processo seletivo da IES administrado pela MANTENEDORA empregadora e a observ�ncia dos preceitos estabelecidos nesta cl�usula e no �Regulamento do Programa de Capacita��o�.
<br><b>Par�grafo quinto</b> � Ter�o direito a requerer e obter do CEBRADE a concess�o de bolsas integrais de estudo, os dependentes legais do PROFESSOR reconhecidos pela Legisla��o do Imposto de Renda, ou que estejam sob a sua guarda judicial e vivam sob sua depend�ncia econ�mica, devidamente comprovada.
<br><b>Par�grafo sexto</b> � Os filhos do PROFESSOR ter�o direito a obter do CEBRADE a concess�o de bolsas de estudo integrais, sem qualquer �nus, desde que n�o tenham vinte e cinco anos completos ou mais na data da efetiva��o da matr�cula no curso superior.
<br><b>Par�grafo s�timo</b> � Caso a MANTENEDORA n�o queira participar do Programa de Amparo Educativo Tempor�rio � PAET, gerenciado pelo CEBRADE, estar� obrigada a conceder bolsas de estudo aos PROFESSORES que trabalham nas Institui��es de Ensino Superior por elas mantidas ou administradas, nas condi��es e termos estabelecidos nesta cl�usula e no Regulamento em anexo.
<br><b>Par�grafo oitavo</b>. Al�m dos casos previstos nesta cl�usula, a MANTENEDORA poder� fornecer outras bolsas de estudos, cujas condi��es ser�o objeto de termo aditivo a ser firmado entre a MANTENEDORA e CEBRADE.

<tr><td class=titulo>16. Assist�ncia m�dico-hospitalar
<tr><td class=campo style="text-align:justify">A MANTENEDORA est� obrigada a assegurar, �s suas expensas, nos limites estabelecidos nesta cl�usula, assist�ncia m�dico-hospitalar a todos os seus PROFESSORES, sendo-lhe facultada a escolha por plano de sa�de, seguro-sa�de ou conv�nios com empresas prestadoras de servi�os m�dico-hospitalares. Poder� ainda prestar a referida assist�ncia diretamente, em se tratando de institui��es que disponham de servi�os de sa�de e hospitais pr�prios ou conveniados. Qualquer que seja a op��o feita, a assist�ncia m�dico-hospitalar deve assegurar as condi��es e os requisitos m�nimos que seguem relacionados:
<br><b>1. Abrang�ncia</b>
<blockquote style="margin-top:0;margin-bottom:0">
	A assist�ncia m�dico-hospitalar deve ser realizada no munic�pio onde funciona o estabelecimento de ensino superior ou onde vive o PROFESSOR, a crit�rio da MANTENEDORA. Em casos de emerg�ncia, dever� haver garantia de atendimento integral em qualquer localidade do Estado de S�o Paulo ou fixa��o, em contrato, de formas de reembolso.
</blockquote>
<b>2. Coberturas m�nimas</b>
<blockquote style="margin-top:0;margin-bottom:0">
	2.1 Quarto para quatro pacientes, no m�ximo.
	<br>2.2 Consultas.
	<br>2.3 Prazo de interna��o de 365 dias por ano (comum e UTI/CTI)
	<br>2.4 Parto, independentemente do estado grav�dico.
	<br>2.5 Mol�stias infecto-contagiosas que exijam interna��o.
	<br>2.6 Exames laboratoriais, ambulatoriais e hospitalares.
</blockquote>
<b>3. Car�ncia</b>
<blockquote style="margin-top:0;margin-bottom:0">
	N�o haver� car�ncia na presta��o dos servi�os m�dicos e laboratoriais.
</blockquote>
<b>4. Professor ingressante</b>
<blockquote style="margin-top:0;margin-bottom:0">
	N�o haver� car�ncia para o PROFESSOR ingressante, independentemente do m�s em que for contratado.
</blockquote>
<b>5. Pagamento</b>
<blockquote style="margin-top:0;margin-bottom:0">
	Caber� ao PROFESSOR o pagamento de 10% do valor da Assist�ncia M�dica, respeitado o disposto nos Par�grafos 1�, 2� e 3�.
</blockquote>
	<b>Par�grafo primeiro</b> � A MANTENEDORA dever� enviar ao Sindicato c�pia do contrato formalizado com a empresa de assist�ncia m�dico�hospitalar ou de seguro sa�de ou de medicina de grupo que comprove o valor pago.
<br><b>Par�grafo segundo</b> � Caso a assist�ncia m�dico-hospitalar vigente na Institui��o venha a sofrer reajuste em virtude de poss�veis modifica��es estabelecidas em legisla��o que abranja o segmento - Lei 9.656, de 03 de junho de 1998 e MP 2.097-39, de 26 de abril de 2001, ou que vierem a ser estabelecidas em lei, ou por mudan�a de empresa prestadora de servi�o, a pedido dos empregados da Institui��o ou por quebra de contrato, unilateralmente, por parte da atual empresa prestadora de servi�o, a MANTENEDORA continuar� a contribuir com o valor mensal vigente at� a data da modifica��o, devendo o PROFESSOR arcar com o valor excedente, que ser� descontado em folha e consignado no comprovante de pagamento, nos termos do artigo 462 da CLT.
<br><b>Par�grafo terceiro</b> � Caso ocorra mudan�a de empresa prestadora de servi�o, por decis�o unilateral da MANTENEDORA, com conseq�ente reajuste no valor vigente, o PROFESSOR estar� isento do pagamento do valor excedente, cabendo � MANTENEDORA prover integralmente a assist�ncia m�dico-hospitalar, sem nenhum �nus para o PROFESSOR.
<br><b>Par�grafo quarto</b> � Para efeito do disposto no Par�grafo primeiro desta cl�usula, caber� � MANTENEDORA remeter a documenta��o comprobat�ria para an�lise e delibera��o da Comiss�o Permanente de Negocia��o.
<br><b>Par�grafo quinto</b> � Fica facultado ao PROFESSOR optar pela presta��o de assist�ncia m�dicohospitalar em uma �nica institui��o de ensino, quando mantiver mais de um v�nculo empregat�cio como PROFESSOR. � necess�rio que o PROFESSOR se manifeste por escrito, com anteced�ncia m�nima de vinte dias, para que a MANTENEDORA possa proceder � suspens�o dos servi�os.
<br><b>Par�grafo sexto</b> � Caso o PROFESSOR mantenha v�nculo empregat�cio com mais de uma Institui��o de Ensino, as MANTENEDORAS, em conjunto, poder�o optar por conceder-lhe um �nico plano de sa�de, pago por elas, em regime de cotiza��o de custos, respeitadas as condi��es estabelecidas nesta cl�usula.
<br><b>Par�grafo s�timo</b> � Mediante pagamento complementar e ades�o facultativa, devidamente documentada, o PROFESSOR poder� optar pela amplia��o dos servi�os de sa�de garantidos nesta Conven��o ou estend�-los a seus dependentes.

<tr><td class=titulo>17. Creches
<tr><td class=campo style="text-align:justify">� obrigat�ria a instala��o de local destinado � guarda de crian�as de at� seis meses, quando a MANTENEDORA mantiver contratada, em jornada integral, pelo menos trinta funcion�rias com idade superior a 16 anos. A manuten��o da creche poder� ser substitu�da pelo pagamento do reembolso creche, nos termos da legisla��o em vigor (artigo 389, Par�grafo 1� da CLT e Portarias MTb n� 3296 de 03/09/86 e n�670, de 27/08/97), ou ainda, a celebra��o de conv�nio com uma entidade reconhecidamente id�nea.

<tr><td class=titulo>18. Sal�rio do professor ingressante na MANTENEDORA
<tr><td class=campo style="text-align:justify">A MANTENEDORA n�o poder� contratar nenhum PROFESSOR por sal�rio inferior ao limite salarial m�nimo dos PROFESSORES mais antigos que possuam o mesmo grau de qualifica��o ou titula��o de quem est� sendo contratado, respeitado o quadro de carreira da MANTENEDORA.
<br><b>Par�grafo �nico</b> � Ao PROFESSOR admitido ap�s 1� de mar�o de 2013 e ap�s 1� de mar�o de 2014, ser�o concedidos os mesmos percentuais de reajustes e aumentos salariais estabelecidos nas cl�usulas �Reajuste salarial em 1� de mar�o de 2013� e �Reajuste salarial em 1� de mar�o de 2014�, respectivamente, desta norma coletiva.

<tr><td class=titulo>19. Readmiss�o do professor
<tr><td class=campo style="text-align:justify">O PROFESSOR que for readmitido at� doze meses ap�s o seu desligamento ficar� desobrigado de firmar contrato de experi�ncia.

<tr><td class=titulo>20. Anota��es na carteira de trabalho
<tr><td class=campo style="text-align:justify">A MANTENEDORA est� obrigada a promover, em quarenta e oito horas, as anota��es nas Carteiras de Trabalho de seus PROFESSORES, ressalvados eventuais prazos mais amplos permitidos por lei.
<br><b>Par�grafo �nico - � obrigat�ria a anota��o na Carteira de Trabalho das mudan�as provocadas por ascens�o ou altera��o de titula��o, decorrentes e previstas em plano de carreira.

<tr><td class=titulo>21. Garantia semestral de sal�rios
<tr><td class=campo style="text-align:justify">Ao PROFESSOR demitido sem justa causa, a MANTENEDORA garantir�:
<blockquote style="margin-top:0;margin-bottom:0">
	a) no primeiro semestre, a partir de 1� de janeiro, os sal�rios integrais at� o dia 30 de junho;
	<br>b) no segundo semestre, os sal�rios integrais at� o dia 31 de dezembro, ressalvado o Par�grafo 4�.
</blockquote>
	<b>Par�grafo primeiro</b> - N�o ter� direito � Garantia Semestral de Sal�rios o PROFESSOR que, na data da comunica��o da dispensa, contar com menos de dezoito meses de servi�o prestado � MANTENEDORA, ressalvado o Par�grafo 4� desta cl�usula.
<br><b>Par�grafo segundo</b> � No caso de demiss�es efetuadas no final do primeiro semestre letivo, para n�o ficar obrigada a pagar ao PROFESSOR os sal�rios do segundo semestre, a MANTENEDORA dever� observar as seguintes disposi��es:
<blockquote style="margin-top:0;margin-bottom:0">
	a) com aviso pr�vio a ser trabalhado, a demiss�o dever� ser formalizada com anteced�ncia m�nima de trinta dias do in�cio das f�rias;
	<br>b) sendo o aviso pr�vio indenizado, a demiss�o dever� ser formalizada at� um dia antes do in�cio das f�rias, ainda que as f�rias tenham seu in�cio programado para o m�s de julho, obedecendo ao que disp�e a cl�usula �F�rias� da presente Conven��o.
</blockquote>
	<b>Par�grafo terceiro</b> - No caso de demiss�es efetuadas no final do ano letivo, para n�o ficar obrigada a pagar ao PROFESSOR os sal�rios do primeiro semestre do ano seguinte, a MANTENEDORA dever� observar as seguintes disposi��es:
<blockquote style="margin-top:0;margin-bottom:0">
	a) com aviso pr�vio a ser trabalhado, a demiss�o dever� ser formalizada com anteced�ncia m�nima de trinta dias do in�cio do recesso escolar;
	<br>b) sendo o aviso pr�vio indenizado, a demiss�o dever� ser formalizada at� um dia antes do in�cio do recesso escolar.
</blockquote>
	<b>Par�grafo quarto</b> - Quando as demiss�es ocorrerem a partir de 16 de outubro, a MANTENEDORA pagar�, independentemente do tempo de servi�o do PROFESSOR, valor correspondente � remunera��o devida at� o dia 18 de janeiro, inclusive, do ano subsequente, respeitado o pagamento m�nimo de trinta dias, a t�tulo de f�rias escolares, para efeito do que define a s�mula 10 do egr�gio TST, ressalvados os contratos de experi�ncia e por prazo determinado, estes �ltimos v�lidos somente nos casos de substitui��o tempor�ria, conforme o disposto na al�nea a) do Par�grafo 2� da cl�usula �Horas extras� da presente Conven��o.
<br><b>Par�grafo quinto</b> � Na vig�ncia da presente Conven��o os PROFESSORES ser�o remunerados a partir da data de in�cio de suas atividades na MANTENEDORA, incluindo o per�odo de planejamento escolar.
<br><b>Par�grafo sexto</b> - Os sal�rios complementares previstos nesta cl�usula ter�o natureza indenizat�ria, n�o integrando, para nenhum efeito legal, o tempo de servi�o do PROFESSOR.
<br><b>Par�grafo s�timo</b> - O aviso pr�vio de trinta dias previsto no artigo 487 da CLT j� est� integrado �s indeniza��es tratadas nesta cl�usula.

<tr><td class=titulo>22. Indeniza��es por dispensa imotivada
<tr><td class=campo style="text-align:justify">�O PROFESSOR demitido sem justa causa, al�m das indeniza��es previstas na cl�usula �Garantia Semestral de Sal�rios� desta Conven��o, ter� direito a receber o valor equivalente a tr�s dias para cada ano trabalhado na MANTENEDORA, nos termos da Lei n� 12.506/2012, sem o limite de tempo de servi�o estabelecido na mesma.
<br><b>Par�grafo primeiro</b> � Caso o PROFESSOR tenha, � data do desligamento, no m�nimo cinquenta anos de idade e conte com pelo menos um ano de servi�o na MANTENEDORA, ter� direito ainda a receber aviso pr�vio adicional indenizado de quinze dias.
<br><b>Par�grafo segundo</b> � N�o ter� direito � indeniza��o assegurada no Par�grafo primeiro o PROFESSOR que na data de admiss�o na MANTENEDORA contar com mais de cinquenta anos de idade.

<tr><td class=titulo>23. Pedido de demiss�o no final de ano letivo
<tr><td class=campo style="text-align:justify">O PROFESSOR que no final do ano letivo comunicar sua demiss�o at� o dia que antecede o in�cio do recesso escolar, ser� dispensado do cumprimento do aviso pr�vio e ter� direito a receber, como indeniza��o, a remunera��o at� o dia 18 de janeiro do ano subseq�ente, independentemente do tempo de servi�o na MANTENEDORA.

<tr><td class=titulo>24. Demiss�o por justa causa
<tr><td class=campo style="text-align:justify">Quando houver demiss�o por justa causa, nos termos do art. 482 da CLT, a MANTENEDORA est� obrigada a determinar na carta-aviso o motivo que deu origem � dispensa. Caso contr�rio, fica descaracterizada a justa causa.

<tr><td class=titulo>25. Multa por atraso na homologa��o
<tr><td class=campo style="text-align:justify">A MANTENEDORA deve pagar as verbas devidas na rescis�o contratual no dia seguinte ao t�rmino do aviso pr�vio, quando trabalhado, ou dez dias ap�s o desligamento, quando houver dispensa do cumprimento de aviso pr�vio. O atraso no pagamento das verbas rescis�rias obrigar� a MANTENEDORA ao pagamento de multa, em favor do PROFESSOR, correspondente a um m�s de sua remunera��o, conforme o disposto no Par�grafo 8� do artigo 477 da CLT.
<br>A partir do vig�simo dia de atraso da homologa��o da rescis�o, a contar da data estabelecida pela legisla��o para o pagamento das verbas rescis�rias, a MATENEDORA estar� obrigada, ainda, a pagar ao PROFESSOR multa di�ria de 0,2% do sal�rio mensal. A MANTENEDORA estar� desobrigada de pagar a referida multa quando o atraso da homologa��o vier a ocorrer, comprovadamente, por motivos alheios a sua vontade.
<br><b>Par�grafo �nico</b> � O Sindicato est� obrigado a fornecer comprovante de comparecimento sempre que a MANTENEDORA se apresentar para homologa��o das rescis�es contratuais e comprovar a convoca��o do PROFESSOR.

<tr><td class=titulo>26. Atestados de afastamento e sal�rios
<tr><td class=campo style="text-align:justify">Sempre que solicitada, a MANTENEDORA dever� fornecer ao PROFESSOR atestado de afastamento e sal�rio (AAS), previsto na legisla��o previdenci�ria.

<tr><td class=titulo>27. Garantia de emprego � gestante
<tr><td class=campo style="text-align:justify">� proibida a dispensa arbitr�ria ou sem justa causa da PROFESSORA gestante, desde o in�cio da gravidez at� sessenta dias ap�s o t�rmino do afastamento legal. O aviso pr�vio come�ar� a contar a partir do t�rmino do per�odo de estabilidade.

<tr><td class=titulo>28. Garantias ao professor com sequelas ocasionadas por doen�as profissionais ou acidente de trabalho
<tr><td class=campo style="text-align:justify">Ser� garantida ao PROFESSOR acidentado no trabalho ou acometido por doen�a profissional a perman�ncia na empresa em fun��o compat�vel com o seu estado f�sico, sem preju�zo na remunera��o antes percebida, desde que, ap�s o acidente ou comprova��o da aquisi��o de doen�a profissional, apresente, cumulativamente, redu��o da capacidade laboral, atestada pelo �rg�o oficial e que se tenha tornado incapaz de exercer a fun��o que anteriormente desempenhava, obrigado, por�m, o PROFESSOR nessa situa��o a participar dos processos de readapta��o e reabilita��o profissional.
<br><b>Par�grafo �nico</b> � O per�odo de estabilidade do PROFESSOR que se encontre participando dos processos de readapta��o e reabilita��o profissional ser� o previsto em lei.

<tr><td class=titulo>29. Estabilidade para portadores de doen�as graves
<tr><td class=campo style="text-align:justify">Fica assegurada, at� alta m�dica, considerada como apto ao trabalho, ou eventual concess�o de aposentadoria por invalidez, estabilidade no emprego aos PROFESSORES acometidos por doen�as graves ou incur�veis e aos PROFESSORES portadores do v�rus HIV que vierem a apresentar qualquer tipo de infec��o ou doen�a oportunista, resultante da patologia de base.
<br><b>Par�grafo �nico</b> � S�o consideradas doen�as graves ou incur�veis, a tuberculose ativa, aliena��o mental, esclerose m�ltipla, neoplasia maligna, cegueira definitiva, hansen�ase, cardiopatia grave, doen�a de Parkinson, paralisia irrevers�vel e incapacitante, espondiloastrose anquilosante, neofropatia grave, estados do Mal de Paget (oste�te deformante) e contamina��o grave por radia��o.

<tr><td class=titulo>30. Garantias ao professor em vias de aposentadoria
<tr><td class=campo style="text-align:justify">Fica assegurado ao PROFESSOR que comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria integral por tempo de servi�o ou da aposentadoria por idade, a garantia de emprego durante o per�odo que faltar at� a aquisi��o do direito.
<br><b>Par�grafo primeiro</b> � A garantia de emprego � devida ao PROFESSOR que estiver contratado pela MANTENEDORA h� pelo menos tr�s anos.
<br><b>Par�grafo segundo</b> � A comprova��o � MANTENEDORA dever� ser feita mediante a apresenta��o de documento que ateste o tempo de servi�o. Este documento dever� ser emitido por pessoa credenciada junto ao �rg�o previdenci�rio. Se o PROFESSOR depender de documenta��o para realiza��o da contagem, ter� um prazo de trinta dias, a contar da data prevista ou marcada para homologa��o da rescis�o contratual. Comprovada a solicita��o de tal documenta��o, os prazos ser�o prorrogados at� que a mesma seja emitida, assegurando-se, nessa situa��o, o pagamento dos sal�rios pelo prazo m�ximo de cento e vinte dias.
<br><b>Par�grafo terceiro</b> � O contrato de trabalho do PROFESSOR s� poder� ser rescindido por m�tuo acordo homologado pelo Sindicato ou pedido de demiss�o.
<br><b>Par�grafo quarto</b> � Havendo acordo formal entre as partes, o PROFESSOR poder� exercer outra fun��o, inerente ao magist�rio, durante o per�odo em que estiver garantido pela estabilidade.
<br><b>Par�grafo quinto</b> � O aviso pr�vio, em caso de demiss�o sem justa causa, integra o per�odo de estabilidade previsto nesta cl�usula.
<br><b>Par�grafo sexto</b> � Para garantir a estabilidade prevista nesta cl�usula, o PROFESSOR dever� encaminhar � MANTENEDORA, dentro da prorroga��o prevista no <br><b>Par�grafo 2�, documenta��o que demonstre a tramita��o do processo que atesta o tempo de servi�o.

<tr><td class=titulo>31. Licen�a � professora adotante
<tr><td class=campo style="text-align:justify">Nos termos da Lei 10421, de 15 de abril de 2002, ser� assegurada licen�a maternidade � PROFESSORA que vier a adotar ou obtiver guarda judicial de crian�as, garantido o emprego no per�odo em que a licen�a for concedida.

<tr><td class=titulo>32. Mudan�a de disciplina
<tr><td class=campo style="text-align:justify">O PROFESSOR n�o poder� ser transferido de uma disciplina para outra, salvo com seu consentimento expresso e por escrito, sob pena de nulidade da referida transfer�ncia.

<tr><td class=titulo>33. Dura��o da hora-aula
<tr><td class=campo style="text-align:justify">A dura��o da hora-aula poder� ser de, no m�ximo, cinquenta minutos.
<br><b>Par�grafo primeiro</b> � Como exce��o ao disposto no caput, a hora-aula poder� ter a dura��o de sessenta minutos nos cursos tecnol�gicos, desde que tenham sido autorizados ou reconhecidos com essa determina��o expressa e cujos PROFESSORES desses cursos tenham sido contratados nessa condi��o.
<br><b>Par�grafo segundo</b> � As MANTENEDORAS de Institui��es de Ensino que possuem cursos tecnol�gicos nas condi��es definidas no Par�grafo 1� desta cl�usula dever�o apresentar � �Comiss�o Permanente de Negocia��o� definida na presente Conven��o, at� o dia 15 de agosto de 2013, a documenta��o de autoriza��o ou reconhecimento do curso com a determina��o expressa de hora aula com dura��o de sessenta minutos sob pena de, em n�o o fazendo, estar sujeita � majora��o do valor do sal�rio-aula de acordo com o que estabelece o Par�grafo 4� desta cl�usula.
<br><b>Par�grafo terceiro</b> � Caso a Comiss�o Permanente de Negocia��o delibere n�o ter havido determina��o expressa do Minist�rio da Educa��o para que a dura��o da hora-aula dos cursos tecnol�gicos seja de sessenta minutos, a MANTENEDORA dever� majorar o sal�rio-aula de acordo com o que estabelece o Par�grafo quarto desta cl�usula.
<br><b>Par�grafo quarto</b> � Em caso de amplia��o da dura��o da hora-aula vigente, respeitado o limite previsto no caput desta cl�usula, a MANTENEDORA dever� acrescer ao sal�rio-aula j� pago, valor proporcional ao acr�scimo do trabalho.

<tr><td class=titulo>34. Carga hor�ria
<tr><td class=campo style="text-align:justify">Quando a MANTENEDORA e o PROFESSOR contratarem carga di�ria de aulas superior aos limites previstos no artigo 318 da CLT, o excedente � carga hor�ria legal ser� remunerado como aula normal, acrescido de DSR, hora-atividade e vantagens pessoais.
<br><b>Par�grafo �nico</b> � Poder� ser flexibilizada a carga hor�ria do PROFESSOR entre jornadas no exerc�cio concomitante de fun��o docente e atividade administrativa, n�o havendo assim pagamento, no intervalo, de horas aulas e sal�rios, se o professor n�o tiver trabalhado no referido intervalo.

<tr><td class=titulo>35. Irredutibilidade de carga hor�ria e de sal�rio
<tr><td class=campo style="text-align:justify">� proibida a redu��o de remunera��o mensal ou de carga hor�ria, ressalvada a ocorr�ncia do disposto nas cl�usulas �Redu��o de carga hor�ria por extin��o de disciplina classe ou turma� e �Redu��o de carga hor�ria por diminui��o do n�mero de alunos matriculados� da presente Conven��o, ou ainda, quando ocorrer iniciativa expressa do PROFESSOR. Em qualquer hip�tese, � obrigat�ria a concord�ncia rec�proca, firmada por escrito.
<br><b>Par�grafo primeiro</b> � N�o havendo concord�ncia rec�proca, a parte que deu origem � redu��o prevista nesta cl�usula arcar� com a responsabilidade da rescis�o contratual.
<br><b>Par�grafo segundo</b> � Outras atividades, ainda que inerentes ao trabalho docente, que n�o sejam as de ministrar aulas, de dura��o tempor�ria e determinada, poder�o ser regulamentadas por contrato entre as partes, contendo a caracteriza��o da atividade, o in�cio e a previs�o do t�rmino.
<br><b>Par�grafo terceiro</b> � A MANTENEDORA n�o poder� reduzir o valor da hora-aula dos contratos de trabalho vigentes, ainda que venha a instituir ou modificar plano de carreira.

<tr><td class=titulo>36. Redu��o de carga hor�ria por extin��o ou supress�o de disciplina, classe ou turma
<tr><td class=campo style="text-align:justify">Ocorrendo supress�o de disciplina, classe ou turma, em virtude de altera��o na estrutura curricular prevista ou autorizada pela legisla��o vigente ou por dispositivo regimental devidamente aprovado por �rg�o colegiado da Institui��o de Ensino, o PROFESSOR da disciplina, classe ou turma dever� ser comunicado da redu��o da sua carga hor�ria, por escrito, com anteced�ncia m�nima de trinta dias do in�cio do per�odo letivo e ter� prioridade para preenchimento de vaga existente em outra classe ou turma ou em outra disciplina para a qual possua habilita��o legal.
<br><b>Par�grafo primeiro</b> � O PROFESSOR dever� manifestar por escrito, no prazo m�ximo de cinco dias ap�s a comunica��o da MANTENEDORA, a n�o-aceita��o da transfer�ncia de disciplina ou de classe ou turma ou da redu��o parcial de sua carga hor�ria. A aus�ncia de manifesta��o do PROFESSOR caracterizar� a sua aceita��o.
<br><b>Par�grafo segundo</b> � Caso o PROFESSOR n�o aceite a transfer�ncia para outra disciplina, classe ou turma ou a redu��o parcial de carga hor�ria, a MANTENEDORA dever� manter a carga hor�ria semanal existente ou proceder � rescis�o do contrato de trabalho, por demiss�o sem justa causa.

<tr><td class=titulo>37. Redu��o de carga hor�ria por diminui��o do n�mero de alunos matriculados
<tr><td class=campo style="text-align:justify">Na ocorr�ncia de diminui��o do n�mero de alunos matriculados que venha a caracterizar a supress�o de turmas, curso ou disciplina, o PROFESSOR do curso em quest�o dever� ser comunicado, por escrito, da redu��o parcial ou total de sua carga hor�ria at� o final da segunda semana de aulas do per�odo letivo.
<br><b>Par�grafo primeiro</b> - O PROFESSOR dever� manifestar, tamb�m por escrito, a aceita��o ou n�o da redu��o parcial de carga hor�ria no prazo m�ximo de cinco dias ap�s a comunica��o da MANTENEDORA. A aus�ncia de manifesta��o do PROFESSOR caracterizar� a sua n�o-aceita��o.
<br><b>Par�grafo segundo</b> - Caso o PROFESSOR aceite a redu��o parcial de carga hor�ria, dever� formalizar documento junto � MANTENEDORA e, em n�o aceitando, a MANTENEDORA dever� proceder � rescis�o do contrato de trabalho, por demiss�o sem justa causa.
<br><b>Par�grafo terceiro</b> - Na hip�tese de rescis�o contratual, por demiss�o sem justa causa, o aviso pr�vio ser� indenizado, estando a MANTENEDORA desobrigada do pagamento do disposto na cl�usula �Garantia Semestral de Sal�rios� da presente Conven��o
<br><b>Par�grafo quarto</b> - N�o ocorrendo redu��o do n�mero de alunos matriculados que venha a caracterizar supress�o do curso, de turma ou de disciplina, a MANTENEDORA que reduzir a carga hor�ria do PROFESSOR estar� sujeita ao disposto na cl�usula �Garantia Semestral de Sal�rios� desta Conven��o quando ocorrer a rescis�o do contrato de trabalho do PROFESSOR.

<tr><td class=titulo>38. Desconto de faltas
<tr><td class=campo style="text-align:justify">Na ocorr�ncia de faltas, a MANTENEDORA poder� descontar do sal�rio do PROFESSOR, no m�ximo, o n�mero de aulas em que o mesmo esteve ausente, o DSR (1/6), a hora-atividade e demais vantagens pessoais proporcionais a estas aulas.
<br><b>Par�grafo �nico</b> - � da compet�ncia e de integral responsabilidade da MANTENEDORA estabelecer mecanismos de controle de faltas e de pontualidade dos PROFESSORES, conforme a legisla��o vigente.

<tr><td class=titulo>39. Abono de faltas por casamento ou luto
<tr><td class=campo style="text-align:justify">N�o ser�o descontadas, no curso de nove dias corridos, as faltas do PROFESSOR, por motivo de gala ou luto, este em decorr�ncia de falecimento de pai, m�e, filho, c�njuge, companheira (o) e dependente juridicamente reconhecido.
<br><b>Par�grafo �nico</b> � N�o ser�o descontadas, no curso de tr�s dias, as faltas do PROFESSOR por motivo de falecimento de sogra, sogro, neto, neta, irm�o ou irm�o.

<tr><td class=titulo>40. Congressos, simp�sios e equivalentes
<tr><td class=campo style="text-align:justify">Os abonos de falta para comparecimento a congressos e simp�sios ser�o concedidos mediante aceita��o por parte da MANTENEDORA, que dever� formalizar por escrito a dispensa do PROFESSOR.
<br><b>Par�grafo �nico</b> - A participa��o do PROFESSOR nos eventos descritos no caput n�o caracterizar� atividade extraordin�ria.

<tr><td class=titulo>41. Janelas
<tr><td class=campo style="text-align:justify">Considera-se janela a aula vaga existente no hor�rio do PROFESSOR entre duas outras aulas ministradas no mesmo turno. O pagamento da janela � obrigat�rio, devendo o PROFESSOR permanecer � disposi��o da MANTENEDORA neste per�odo, ressalvada a aceita��o pelo PROFESSOR, atrav�s de acordo formalizado entre as partes antes do in�cio das aulas, quando as janelas n�o ser�o pagas.
<br><b>Par�grafo �nico</b> - Ocorrendo a hip�tese da ressalva supra e caso o PROFESSOR seja solicitado esporadicamente a ministrar aulas ou a desenvolver qualquer outra atividade inerente ao magist�rio, no hor�rio de janelas n�o-pagas, essas atividades ser�o remuneradas como aulas extras, com adicional de 100%.

<tr><td class=titulo>42. F�rias
<tr><td class=campo style="text-align:justify">As f�rias anuais dos PROFESSORES ser�o coletivas, com dura��o de trinta dias corridos e gozados em julho de 2013 e julho de 2014. Qualquer altera��o dever� ser aprovada por �rg�o competente, conforme o estabelecido em Estatuto ou Regimento e dever� constar do calend�rio escolar, obrigatoriamente divulgado aos PROFESSORES at� o in�cio de cada per�odo letivo e enviado ao Sindicato.
<br><b>Par�grafo primeiro</b> � A MANTENEDORA est� obrigada a pagar o sal�rio das f�rias e o abono constitucional de 1/3 at� quarenta e oito horas antes do in�cio das f�rias.
<br><b>Par�grafo segundo</b> � As f�rias n�o poder�o ser iniciadas aos domingos, feriados, dias de compensa��o do descanso semanal remunerado e nem aos s�bados, quando estes n�o forem dias normais de aula.
<br><b>Par�grafo terceiro</b> � Tamb�m ter� direito �s f�rias coletivas de trinta dias corridos nos per�odos estabelecidos no caput, O PROFESSOR que, al�m de ministrar aulas, tenha cargo de confian�a ou exer�a outras atividades na MANTENEDORA.
	<br>Caso o exerc�cio da atividade administrativa impossibilite a concess�o de f�rias nos termos do caput, as f�rias anuais desse PROFESSOR poder�o ser gozadas em dois per�odos, um deles obrigatoriamente no m�s de julho de cada ano.
<br><b>Par�grafo quarto</b> � Na hip�tese da divis�o das f�rias anuais do PROFESSOR nos termos do Par�grafo anterior, um dos per�odos n�o poder� ser inferior a dez dias, sendo proibido o exerc�cio de qualquer atividade nesses per�odos.

<tr><td class=titulo>43. Recesso escolar
<tr><td class=campo style="text-align:justify">O recesso escolar anual � obrigat�rio e tem dura��o de trinta dias corridos, gozados preferencialmente no m�s de janeiro de cada ano.
<br>Durante o recesso escolar anual que n�o pode, de maneira alguma, coincidir com o per�odo definido para as f�rias coletivas do ano respectivo, o PROFESSOR n�o poder� ser convocado para nenhum trabalho.
<br><b>Par�grafo primeiro</b> � Na vig�ncia da presente Conven��o, as institui��es cujos calend�rios escolares, determinados pelo �rg�o competente conforme o estabelecido em Estatuto ou Regimento, n�o observarem o determinado pelo caput para o recesso escolar anual dos PROFESSORES, poder�o conced�-lo em um per�odo de no m�nimo vinte dias corridos e em no m�ximo mais tr�s per�odos compostos por dias normais de aula e consecutivos, desde que observem as seguintes condi��es:
<blockquote style="margin-top:0;margin-bottom:0">
	a) vinte dias corridos em janeiro de 2014 e os dois ou tr�s per�odos compostos por dias normais de aula e consecutivos, obrigatoriamente no per�odo compreendido entre mar�o de 2013 e fevereiro de 2014.
	<br>b) vinte dias corridos em janeiro de 2015 e os dois ou tr�s per�odos compostos por dias letivos e consecutivos, obrigatoriamente no per�odo compreendido entre mar�o de 2014 e fevereiro de 2015.
</blockquote>
	<b>Par�grafo segundo</b> � No caso dos calend�rios escolares preverem a divis�o do recesso escolar dos PROFESSORES, os per�odos definidos na conformidade do Par�grafo primeiro n�o poder�o ser iniciados aos domingos, feriados, dias de compensa��o do descanso semanal remunerado e nem aos s�bados, quando estes n�o forem dias normais de aulas.
<br><b>Par�grafo terceiro</b> � As Institui��es cujas atividades n�o possam ser interrompidas, tais como aquelas desenvolvidas em hospital, cl�nica, laborat�rio de an�lise, escrit�rios experimentais, pesquisas, dentre outros, ou que ministrem cursos em que sejam utilizadas instala��es espec�ficas ou que prestem atendimento � comunidade que n�o pode ser suspenso, poder�o conceder aos PROFESSORES o recesso escolar anual definido no caput de maneira escalonada ao longo de cada ano.
<br><b>Par�grafo quarto</b> � Os calend�rios escolares que definir�o os per�odos de recesso escolar dos PROFESSORES ser�o obrigatoriamente divulgados aos PROFESSORES at� o in�cio de cada per�odo letivo e enviados ao Sindicato.

<tr><td class=titulo>44. Licen�a sem remunera��o
<tr><td class=campo style="text-align:justify">O PROFESSOR com mais de cinco anos ininterruptos de servi�o na MANTENEDORA ter� direito a licenciar-se, sem remunera��o, por um per�odo m�ximo de dois anos, n�o sendo este per�odo de afastamento computado para contagem de tempo de servi�o ou para qualquer outro efeito, inclusive legal.
<br><b>Par�grafo primeiro</b> - A licen�a ou sua prorroga��o dever� ser comunicada por escrito, � MANTENEDORA, com anteced�ncia m�nima de noventa dias do per�odo letivo, devendo especificar as datas de in�cio e t�rmino do afastamento. A licen�a s� ter� in�cio a partir da data expressa no comunicado, mantendo-se, at� a�, todas as vantagens contratuais. A inten��o de retorno do PROFESSOR � atividade dever� ser comunicada � MANTENEDORA, no m�nimo, sessenta dias antes do t�rmino do afastamento.
<br><b>Par�grafo segundo</b> - O t�rmino do afastamento dever� coincidir com o in�cio do per�odo letivo.
<br><b>Par�grafo terceiro</b> - O PROFESSOR que tenha ou exer�a cargo de confian�a dever�, junto com o comunicado de licen�a, solicitar seu desligamento do cargo a partir do in�cio do per�odo de licen�a.
<br><b>Par�grafo quarto</b> - Considera-se demission�rio o PROFESSOR que, ao t�rmino do afastamento, n�o retornar �s atividades docentes.
<br><b>Par�grafo quinto</b> - Ocorrendo a dispensa sem justa causa ao t�rmino da licen�a, o PROFESSOR n�o ter� direito � �Garantia Semestral de Sal�rios�, prevista na presente Conven��o.

<tr><td class=titulo>45. Licen�a paternidade
<tr><td class=campo style="text-align:justify">A licen�a paternidade ter� dura��o de cinco dias.

<tr><td class=titulo>46. Uniformes
<tr><td class=campo style="text-align:justify">A MANTENEDORA dever� fornecer gratuitamente dois uniformes por ano, quando o seu uso for exigido.

<tr><td class=titulo>47. Atestados m�dicos e abono de faltas
<tr><td class=campo style="text-align:justify">A MANTENEDORA est� obrigada a abonar as faltas dos PROFESSORES, mediante a apresenta��o de atestados m�dicos ou odontol�gicos.

<tr><td class=titulo>48. Quadro de avisos
<tr><td class=campo style="text-align:justify">A MANTENEDORA dever� colocar, nas salas de professores, quadro de aviso � disposi��o do Sindicato para fixa��o de comunicados de interesse da categoria, sendo vedada a divulga��o de mat�ria pol�tico-partid�ria ou ofensiva a quem quer que seja.
<br><b>Par�grafo �nico</b> � O dirigente sindical ter� livre acesso � sala dos PROFESSORES, no hor�rio de intervalo das aulas, para atualiza��o do material divulgado no quadro de avisos, uma �nica vez em cada m�s.

<tr><td class=titulo>49. Delegado representante
<tr><td class=campo style="text-align:justify">A MANTENEDORA assegurar� a elei��o de um Delegado Representante para cada Institui��o de Ensino Superior mantida, com mandato de um ano, que ter� a garantia de emprego e sal�rios a partir da inscri��o de sua candidatura at� o t�rmino do semestre letivo em que sua gest�o se encerrar.
<br><b>Par�grafo primeiro</b> � A elei��o dos Delegados Representantes ser� realizada pelo Sindicato na Institui��o de Ensino Superior mantida, por voto direto e secreto. � exigido qu�rum de 50% mais um do corpo docente da unidade onde a elei��o ocorrer.
<br><b>Par�grafo segundo</b> � O Sindicato comunicar� a elei��o � MANTENEDORA, com a rela��o dos candidatos inscritos, com anteced�ncia m�nima de sete dias corridos da data da elei��o. Nenhum candidato poder� ser demitido a partir da data da comunica��o at� o t�rmino da apura��o.
<br><b>Par�grafo terceiro</b> � � condi��o necess�ria que os candidatos sejam filiados ao Sindicato e que tenham, � data da elei��o, pelo menos um ano de servi�o na MANTENEDORA.

<tr><td class=titulo>50. Assembleias sindicais
<tr><td class=campo style="text-align:justify">Todo PROFESSOR ter� direito a abono de faltas para o comparecimento a assembleias da categoria.
<br><b>Par�grafo primeiro</b> - Na vig�ncia desta Conven��o, os abonos est�o limitados a dois s�bados e mais dois dias �teis para cada per�odo compreendido entre o m�s de mar�o e o m�s de fevereiro do ano subsequente. As duas assembleias realizadas durante os dias �teis dever�o ocorrer em per�odos distintos.
<br><b>Par�grafo segundo</b> - O Sindicato ou a FEPESP dever� informar ao SEMESP ou � MANTENEDORA, por escrito, com anteced�ncia m�nima de quinze dias corridos. Na comunica��o dever�o constar a data e o hor�rio da assembleia.
<br><b>Par�grafo terceiro</b> - Os dirigentes sindicais n�o est�o sujeitos ao limite previsto no Par�grafo 1� desta cl�usula. As aus�ncias decorrentes do comparecimento �s assembleias de suas entidades ser�o abonadas mediante pr�via comunica��o formal � MANTENEDORA.
<br><b>Par�grafo quarto</b> - A MANTENEDORA poder� exigir dos PROFESSORES e do dirigente sindical atestado emitido pelo Sindicato ou pela FEPESP que comprove o seu comparecimento � assembleia.

<tr><td class=titulo>51. Congresso do Sinpro
<tr><td class=campo style="text-align:justify">Em cada ano de vig�ncia desta Conven��o, o Sindicato promover� um evento de natureza pol�tica ou pedag�gica (congresso ou jornada). A MANTENEDORA abonar� as aus�ncias de seus PROFESSORES que participarem do evento, nos seguintes limites:
<blockquote style="margin-top:0;margin-bottom:0">
	a) na unidade de ensino que tenha at� quarenta e nove PROFESSORES ser� garantido o abono a um PROFESSOR;
	<br>b) na unidade de ensino que tenha entre cinquenta e noventa e nove PROFESSORES ser� garantido o abono a dois PROFESSORES;
	<br>c) na unidade de ensino que tenha mais de cem PROFESSORES ser� garantido o abono a tr�s PROFESSORES.
</blockquote>
	Tais faltas, limitadas ao m�ximo em dois dias �teis al�m do s�bado, em cada evento, ser�o abonadas mediante a apresenta��o de atestado de comparecimento fornecido pelo Sindicato. O PROFESSOR dever� repor as aulas que, por ventura, sejam necess�rias para complementa��o das horas letivas m�nimas exigidas pela legisla��o.

<tr><td class=titulo>52. Rela��o nominal
<tr><td class=campo style="text-align:justify">Na vig�ncia desta Conven��o, obriga-se a MANTENEDORA a encaminhar ao Sindicato, at� o final do m�s de junho de cada ano, a rela��o nominal dos PROFESSORES que integram seu quadro de funcion�rios, acompanhada do valor do sal�rio mensal e das guias das contribui��es sindical e assistencial.

<tr><td class=titulo>53. Acordos internos - cl�usulas mais favor�veis
<tr><td class=campo style="text-align:justify">Ficam assegurados os direitos mais favor�veis decorrentes de acordos internos ou de acordos coletivos de trabalho celebrados entre a MANTENEDORA e o Sindicato.

<tr><td class=titulo>54. Comiss�o Permanente de Negocia��o
<tr><td class=campo style="text-align:justify">Fica mantida a Comiss�o Permanente de Negocia��o constitu�da de forma parit�ria, por tr�s representantes das entidades sindicais (profissional e econ�mica), com o objetivo de:
<blockquote style="margin-top:0;margin-bottom:0">
	a) fiscalizar o cumprimento das cl�usulas vigentes;
	<br>b) elucidar eventuais diverg�ncias de interpreta��o das cl�usulas desta Conven��o;
	<br>c) discutir quest�es n�o contempladas na presente Conven��o.
	<br>d) deliberar no prazo m�ximo de trinta dias a contar da data da solicita��o protocolizada no SEMESP, sobre modifica��o de pagamento da assist�ncia m�dico-hospitalar, conforme os Par�grafos 1� e 3� da cl�usula �Assist�ncia M�dico Hospitalar� desta Conven��o e sobre o valor da remunera��o da hora-aula, conforme o Par�grafo 2� da cl�usula �Dura��o da hora-aula� desta Conven��o.
	<br>e) criar subs�dios para a Comiss�o de Tratativas Salariais, atrav�s da elabora��o de documentos, para a defini��o das fun��es/atividades e o regime de trabalho dos PROFESSORES.
</blockquote>
	<b>Par�grafo primeiro</b> - As entidades sindicais componentes da Comiss�o Permanente de Negocia��o indicar�o seus representantes, no prazo m�ximo de trinta dias corridos, a contar da assinatura desta Conven��o.
<br><b>Par�grafo segundo</b> - A Comiss�o Permanente de Negocia��o dever� reunir-se mensalmente, no d�cimo dia �til, �s 15 horas, alternadamente nas sedes das entidades sindicais que a comp�em. No caso espec�fico do item �d� do caput, dever� haver convoca��o espec�fica feita pelo SEMESP.

<tr><td class=titulo>55. Disposi��es transit�rias
<tr><td class=campo style="text-align:justify">Fica mantida a Comiss�o de Aprimoramento das Rela��es de Trabalho, composta de forma parit�ria, por quatro membros de cada uma das categorias econ�mica e profissional, indicados, respectivamente, pelo SEMESP e/ou SEMESP/SJ RIO PRETO e pelas Federa��es � FEPESP, FEPAAE e FETEE, com o objetivo de apresentar proposta de regulamenta��o dos seguintes temas: 
<blockquote style="margin-top:0;margin-bottom:0">
	rela��es de trabalho envolvendo a defini��o de atividade docente e aplica��es de novas tecnologias (hora tecnol�gica); 
	<br>ensino a dist�ncia; 
	<br>defini��o e regulamenta��o da atividade de tutor de cursos � dist�ncia e semipresenciais, cursos semipresenciais e cursos modulares e sequenciais; 
	<br>planos de carreira das Institui��es de Ensino Superior privadas; 
	<br>bolsas de estudos e plano de sa�de, adicional de titula��o, aprimoramento acad�mico, n�mero m�ximo de alunos em sala de aula, complementa��o do aux�lio doen�a, piso salarial e aux�lio creche.
</blockquote>
	<b>Par�grafo primeiro</b> � Caso o percentual definido pela m�dia aritm�tica dos �ndices inflacion�rios do per�odo compreendido entre 1� de mar�o de 2013 e 28 de fevereiro de 2014, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV), ultrapasse o limite de 6,5%, o pagamento da diferen�a entre aquela m�dia aritm�tica e 6,5% ser� negociado pela Comiss�o de Aprimoramento das Rela��es de Trabalho, no prazo de noventa dias, a contar de 1� de abril de 2014, sendo certo que, para base de c�lculo de mar�o de 2015, est� garantida a f�rmula de c�lculo definida no Par�grafo primeiro da cl�usula Reajuste Salarial em 1� de mar�o de 2014, sem o limite estabelecido neste Par�grafo.
<br><b>Par�grafo segundo</b> � O regimento de funcionamento da Comiss�o de Aprimoramento das Rela��es de Trabalho, que poder� prever mecanismos de concilia��o e/ou media��o, ser� definido na primeira reuni�o a ser convocada por qualquer uma das partes envolvidas.
<br><b>Par�grafo terceiro</b> � Os estudos, relat�rios e delibera��es da �Comiss�o de Aprimoramento das Rela��es do Trabalho�, ser�o submetidos �s delibera��es das Assembleias convocadas pelas respectivas entidades sindicais, at� 31 de mar�o de 2014 e, uma vez aprovadas, aditadas a esta Conven��o.

<tr><td class=titulo>56. Foro Conciliat�rio para Solu��o de Conflitos Coletivos
<tr><td class=campo style="text-align:justify">Fica mantida a exist�ncia do Foro Conciliat�rio que tem como objetivo procurar resolver quest�es referentes ao n�o cumprimento de normas estabelecidas na presente Conven��o e eventuais diverg�ncias trabalhistas existentes entre a MANTENEDORA e seus PROFESSORES.
<br><b>Par�grafo primeiro</b> - O Foro ser� composto por membros do SEMESP e do Sindicato. As reuni�es dever�o contar, tamb�m, com as partes em conflito que, se assim o desejarem, poder�o delegar representantes para substitu�-las e/ou serem assistidas por advogados.
<br><b>Par�grafo segundo</b> - O SEMESP e o Sindicato dever�o indicar os seus representantes no Foro num prazo de trinta dias a contar da assinatura desta Conven��o.
<br><b>Par�grafo terceiro</b> - Cada se��o do Foro ser� realizada no prazo m�ximo de quinze dias a contar da solicita��o formal e obrigat�ria de qualquer uma das entidades que o comp�em, devendo constar na solicita��o a data, o local e o hor�rio em que a mesma dever� se realizar. O n�o comparecimento de qualquer uma das partes acarretar� no encerramento imediato das negocia��es.
<br><b>Par�grafo quarto</b> - Nenhuma das partes envolvidas ingressar� com a��o na Justi�a do Trabalho durante as negocia��es de entendimento.
<br><b>Par�grafo quinto</b> - Na aus�ncia de solu��o do conflito ou na hip�tese de n�o-comparecimento de qualquer uma das partes, a comiss�o respons�vel pelo Foro fornecer� certid�o atestando o encerramento da negocia��o.
<br><b>Par�grafo sexto</b> - Na hip�tese de sucesso das negocia��es, a crit�rio do Foro, a MANTENEDORA ficar� desobrigada de arcar com a multa definida na cl�usula �Multa por descumprimento da Conven��o�.
<br><b>Par�grafo s�timo</b> - As decis�es do Foro ter�o efic�cia legal entre as partes acordantes. O descumprimento das decis�es assumidas gerar� multa a ser estabelecida no Foro, independentemente daquelas j� estabelecidas nesta Conven��o.
<br><b>Par�grafo oitavo</b> � Na hip�tese de incapacidade econ�mico-financeira das MANTENEDORAS, os casos ser�o remetidos para an�lise e delibera��o deste foro.

<tr><td class=titulo>57. Multa por descumprimento da Conven��o
<tr><td class=campo style="text-align:justify">O descumprimento desta Conven��o obrigar� a MANTENEDORA ao pagamento de multa correspondente a 1% do sal�rio do PROFESSOR, para cada uma das cl�usulas n�o cumpridas, acrescidas de juros, a cada PROFESSOR prejudicado, � exce��o da cl�usula �Participa��o nos lucros ou resultados ou abono especial�, cujo descumprimento obrigar� a MANTENDORA ao pagamento a cada PROFESSOR de multa espec�fica no valor de 24% do seu sal�rio bruto do m�s de outubro de 2014.
<br><b>Par�grafo �nico</b> � A MANTENEDORA est� desobrigada de arcar com a multa prevista no caput, caso a cl�usula descumprida j� estabele�a uma multa pelo seu n�o cumprimento.

<tr><td class=titulo>58. Contribui��o Assistencial
<tr><td class=campo style="text-align:justify"><i>Cl�usula v�lida para os seguintes sindicatos: Sinpro ABC � Sinpro Ara�atuba e Birigui � Sinpro Bauru e regi�o � Sinpro Campinas e regi�o � Sinteee Franca � Sinpro Guapira � Sinpro Guarulhos � Sinpro Jacare� � Sinpro Ja� � Sinpro Jundia� �Sintee Lins �Sinpro Osasco e regi�o � Sintee Presidente Prudente e regi�o � Sinpaae Ribeir�o Preto e regi�o � Sinpro Santos e regi�o � Sinpro S�o Carlos �Sinpro Sorocaba e regi�o � Sinpro Taubat� �Sinpro Unicidades �Sinpro Vales � Sinpro Valinhos-Vinhedo</i>
<tr><td class=campo style="text-align:justify">Obriga-se a MANTENEDORA a promover o desconto da contribui��o assistencial, na folha de pagamento de seus PROFESSORES, sindicalizados e/ou filiados ou n�o, para recolhimento em favor do Sindicato profissional, conforme base territorial definida no MTE, em conta especial, na import�ncia deliberada pelas respectivas Assembleias Gerais, se observados os Par�grafos abaixo.
<br><b>Par�grafo primeiro</b> � Fica assegurado ao PROFESSOR o direito de oposi��o � cobran�a da contribui��o assistencial, a ser exercido, sem qualquer v�cio de vontade, em trinta dias ap�s a entrada em vigor da presente Conven��o Coletiva, com o dep�sito perante o Minist�rio do Trabalho e Emprego, no 1� ano de sua vig�ncia e, no 2� ano de vig�ncia, no per�odo compreendido entre 1� e 31 de mar�o de 2014, a ser exercido de modo individual, pessoalmente ou por meio de carta registrada encaminhada ao Sindicato profissional, com c�pia � entidade Mantenedora.
<br><b>Par�grafo segundo</b> � O recolhimento da contribui��o assistencial ser� realizado obrigatoriamente pela pr�pria MANTENEDORA, at� o 10� dia dos meses subsequentes aos descontos, em guias pr�prias, fornecidas pelo Sindicato da categoria profissional.
<br><b>Par�grafo terceiro</b> - Os Sindicatos representantes das categorias patronal e profissional ficam obrigados a informar, respectivamente, a cada categoria representada (atrav�s de publica��o em site da entidade na internet, publica��o de edital em jornal de ampla circula��o na localidade, no quadro de avisos dos empregados na institui��o e outros meios eficazes), incluindo informa��es sobre a cobran�a das referidas contribui��es e as condi��es para o exerc�cio de oposi��o, nos seguintes prazos: em cinco dias �teis imediatamente ap�s assinatura da Conven��o Coletiva, no 1� ano de vig�ncia e at� o dia 5 de mar�o de 2014, no segundo ano de vig�ncia.
<br><b>Par�grafo quarto</b> - A Assembleia para autoriza��o da contribui��o assistencial dever� atender aos seguintes requisitos: 
<blockquote style="margin-top:0;margin-bottom:0">
	1) o edital de convoca��o da Assembleia Geral dever� ter ampla divulga��o, com a publica��o em jornais de grande circula��o, especialmente convocada para a aprova��o da contribui��o assistencial, garantindo-se o acesso a todos os trabalhadores, s�cios e n�o s�cios; 
	<br>2) realiza��o em local e hor�rio que facilitem a presen�a dos trabalhadores; 
	<br>3) observa��o dos princ�pios da proporcionalidade e razoabilidade, para fixa��o do valor da contribui��o assistencial, sendo considerado razo�vel o valor da contribui��o correspondente at� 1% ao m�s, n�o cumulativa, at� 5% por ano de vig�ncia da norma contratual coletiva, calculada sobre o valor do sal�rio bruto reajustado por ocasi�o de cada norma coletiva da categoria.
</blockquote>
	<b>Par�grafo quinto</b> � Para que a contribui��o assistencial possa ser pleiteada pelo Sindicato da categoria profissional, o SEMESP dever� receber o edital de convoca��o e a ata que deliberou sobre a referida contribui��o, no prazo de cinco dias �teis ap�s a sua realiza��o e anteriormente a inclus�o da presente norma no Sistema Mediador.
<br><b>Par�grafo sexto</b> � As Federa��es representativas dos Sindicatos profissionais dever�o encaminhar ao SEMESP, antes de qualquer assinatura de conven��o coletiva, c�pias de termos de ajustamento de conduta assinados com o Minist�rio P�blico ou decis�es judiciais acerca de contribui��o assistencial.
<br><b>Par�grafo s�timo</b> - O descumprimento de qualquer dos Par�grafos anteriores acarretar� multa di�ria de R$ 1.000,00, nos termos do art. 461, 4� do C�digo de Processo Civil at� comprova��o de regulariza��o da conduta, sendo revertidos os valores ao FAT � Fundo de Amparo ao Trabalhador.
<br><b>Par�grafo oitavo</b> � Fica expressamente ressalvado que a presente cl�usula n�o prejudica e nem beneficia terceiros que possuam a��o judicial ou termo de ajustamento de conduta com entendimento diverso do acima estabelecido, nem a defesa dos direitos individuais de cada trabalhador que se sentir prejudicado.

<tr><td class=campo style="text-align:justify">Conven��o assinada em 15 de agosto de 2013, com efeitos retroativos a 1� de mar�o de 2013.

</table>
<%
'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>
<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="../images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>