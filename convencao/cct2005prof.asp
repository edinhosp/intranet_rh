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
<title>Conven��o Coletiva 2005 - Professores</title>
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
<tr><td class=titulo align="center">CONVEN��O COLETIVA DE TRABALHO 2005
<tr><td class=titulo align="center">SEMESP
<tr><td class=titulo align="center">PROFESSORES 
<tr><td class=campo style="text-align:justify">Entre as partes, de um lado, o Sindicato dos Professores de S�o Jos� do Rio Preto - SINPRO S�o Paulo; SINPRO de OSASCO; SINPRO de Santos e Regi�o; SINPRO de Jundia�; SINPRO de Guarulhos; ee a Federa��o dos Professores do Estado de S�o Paulo � FEPESP, entidades com bases territoriais e representatividades fixadas nas respectivas Cartas Sindicais e no que estabelece o inciso I do artigo 8� da Constitui��o Federal e de outro, o Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de S�o Paulo � SEMESP e SEMESP S�o Jos� do Rio Preto, com representatividade fixada em seus registros sindicais, ao final assinados por seus representantes legais, devidamente autorizados pelas competentes Assembl�ias Gerais das respectivas categorias, fica estabelecida, nos termos do artigo 611 e seguintes da Consolida��o das Leis do Trabalho e do artigo 8�, inciso VI da Constitui��o Federal, a presente CONVEN��O COLETIVA DE TRABALHO:

<tr><td class=titulo>1. ABRANG�NCIA
<tr><td class=campo style="text-align:justify">Esta Conven��o abrange a categoria econ�mica dos estabelecimentos particulares de ensino superior no Estado de S�o Paulo, aqui designados como <b>MANTENEDORA</b> e a categoria profissional diferenciada dos professores, aqui designada simplesmente como PROFESSOR.
<br><b>Par�grafo �nico</b> � A categoria dos PROFESSORES abrange todos aqueles que exercem a atividade docente, independentemente da denomina��o sob a qual a fun��o for exercida. Considera-se atividade docente a fun��o de ministrar aulas.

<tr><td class=titulo>2. Dura��o
<tr><td class=campo style="text-align:justify">Esta Conven��o Coletiva de Trabalho ter� dura��o de dois anos, com vig�ncia de 1� de mar�o de 2005 a 28 de fevereiro de 2007.
<br><b>Par�grafo �nico</b> � As cl�usulas poder�o ser reexaminadas na pr�xima data-base em virtude de problemas surgidos na sua aplica��o ou do surgimento de normas legais a elas pertinentes.

<tr><td class=titulo>3. Reajuste Salarial em 1� de maio de 2005
<tr><td class=campo style="text-align:justify">Sobre os sal�rios devidos em 1� de junho de 2004 ser� aplicado, a partir de 1� de maio de 2005, um reajuste salarial de 7,66% (sete virgula sessenta e seis por cento), observado o estabelecido na cl�usula 4� da presente Conven��o.
<br><b>Par�grafo primeiro</b> - Fica estabelecido que o sal�rio de 1� de maio de 2005, reajustado pelo �ndice definido nesta cl�usula, servir� como base de c�lculo para a data base de 1� de mar�o de 2006.
<br><b>Par�grafo segundo</b> � Eventuais diferen�as salariais resultantes da aplica��o da presente norma coletiva, at� a data da sua assinatura, dever�o ser pagas at� o dia 15 de setembro de 2005, sem incid�ncia de multa convencional.

<tr><td class=titulo>4. REAJUSTE SALARIAL em 1� de mar�o de 2006
<tr><td class=campo style="text-align:justify">Em 1� de mar�o de 2006, as MANTENEDORAS dever�o aplicar sobre os sal�rios devidos em 1� de maio de 2005, o percentual definido pela m�dia aritm�tica dos �ndices inflacion�rios do per�odo compreendido entre 1� de mar�o de 2005 e 28 de fevereiro de 2006, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV).
<br><b>Par�grafo primeiro</b> � Se a m�dia aritm�tica dos �ndices inflacion�rios definida no caput superar 9,99% (nove virgula noventa e nove por cento), as MANTENEDORAS dever�o aplicar, em 1� de mar�o de 2006, sobre os sal�rios devidos em 1� de maio de 2005, o reajuste de 9,99% (nove virgula noventa e nove por cento). O SEMESP, o SINPRO e a FEPESP definir�o, em processo de negocia��o salarial, at� o prazo m�ximo de 30 de abril de 2006, a forma de pagamento da parcela excedente a 9,99%.
<br><b>Par�grafo segundo</b> � O SEMESP, o SINPRO e a FEPESP comprometem-se a divulgar, em comunicado conjunto, at� 20 de mar�o de 2006, o percentual de reajuste salarial calculado pela f�rmula definida no caput, bem como a forma de pagamento da parcela excedente a 9,99%.
<br><b>Par�grafo terceiro</b> � A base de c�lculo para a data-base de 1� de mar�o de 2007 ser� constitu�da pelos sal�rios devidos em 1� de maio de 2005, reajustados em 2006 pela m�dia aritm�tica dos �ndices inflacion�rios do per�odo compreendido entre 1� de mar�o de 2005 e 28 de fevereiro de 2006, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV).

<tr><td class=titulo>5. COMPENSA��ES SALARIAIS
<tr><td class=campo style="text-align:justify">Para 2005 ser� permitida a compensa��o de eventuais antecipa��es salariais concedidas no per�odo de vig�ncia da Conven��o de 2004. Relativamente � Conven��o de 2006, ser� permitida a compensa��o de eventuais antecipa��es salariais concedidas no per�odo de vig�ncia da Conven��o de 2005.
<br><b>Par�grafo �nico</b> � Excetuam-se em ambos os casos aquelas que decorrerem de promo��es, transfer�ncias, ascens�o em plano de carreira e aqueles reajustes concedidos com cl�usula expressa de n�o compensa��o.

<tr><td class=titulo>6. SAL�RIO DO <b>PROFESSOR</b> INGRESSANTE NA MANTENEDORA
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> n�o poder� contratar nenhum <b>PROFESSOR</b> por sal�rio inferior ao limite salarial m�nimo dos PROFESSORES mais antigos que possuam o mesmo grau de qualifica��o ou titula��o de quem est� sendo contratado, respeitado o quadro de carreira da MANTENEDORA.
<br><b>Par�grafo �nico</b> � Ao <b>PROFESSOR</b> admitido ap�s 1� de mar�o de 2005 e ap�s 1� de mar�o de 2006, respectivamente, ser�o concedidos os mesmos percentuais de reajustes e aumentos salariais estabelecidos nesta norma coletiva.

<tr><td class=titulo>7. COMPROVANTE DE PAGAMENTO
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> dever� fornecer ao PROFESSOR, mensalmente, comprovante de pagamento, devendo estar discriminados: a)identifica��o da <b>MANTENEDORA</b> e do estabelecimento de ensino; b)a identifica��o do PROFESSOR; c)a denomina��o da categoria, se houver faixas salariais diferenciadas; d)o valor da hora-aula; e)a carga hor�ria semanal; f)a hora-atividade; g)outros eventuais adicionais; h)o descanso semanal remunerado; i)as horas extras realizadas; j)o valor do recolhimento do FGTS; l)o desconto previdenci�rio; m)outros descontos.

<tr><td class=titulo>8. HORA-ATIVIDADE
<tr><td class=campo style="text-align:justify">Fica mantido o adicional de 5% (cinco por cento) a t�tulo de hora-atividade, destinado exclusivamente ao pagamento do tempo gasto pelo PROFESSOR, fora do estabelecimento de ensino, na prepara��o de aulas, provas e exerc�cios, bem como na corre��o dos mesmos.

<tr><td class=titulo>9. ADICIONAL NOTURNO
<tr><td class=campo style="text-align:justify">O trabalho noturno deve ser pago nas atividades realizadas ap�s as 22 horas e corresponde a 25% (vinte e cinco por cento) do valor da hora-aula.

<tr><td class=titulo>10. HORAS EXTRAS
<tr><td class=campo style="text-align:justify">Considera-se atividade extra todo trabalho desenvolvido em hor�rio diferente daquele habitualmente realizado na semana. As atividades extras devem ser pagas com adicional de 100% (cem por cento).
<br><b>Par�grafo primeiro</b> - N�o � considerada atividade extra a participa��o em cursos de capacita��o e aperfei�oamento docente, desde que aceita livremente pelo PROFESSOR.
<br><b>Par�grafo segundo</b> - Ser�o pagas apenas como aulas normais, acrescidas do DSR e da hora-atividade, aquelas que forem adicionadas provisoriamente � carga hor�ria habitual, decorrentes:
<blockquote style="margin-top:0;margin-bottom:0">a) da substitui��o tempor�ria de um outro PROFESSOR, com dura��o predeterminada, decorrente de licen�a m�dica, maternidade ou para estudos. Nestes casos, a substitui��o dever� ser formalizada atrav�s de documento firmado entre a <b>MANTENEDORA</b> e o <b>PROFESSOR</b> que aceitar realiz�-la;
<br>b) de substitui��es eventuais de faltas de <b>PROFESSOR</b> respons�vel, desde que aceitas livremente pelo <b>PROFESSOR</b> substituto;
<br>c) de reposi��o de eventuais faltas que foram descontadas dos sal�rios nos meses em que ocorreram;
<br>d) da realiza��o de cursos eventuais ou de curta dura��o, inclusive cursos de depend�ncia, e aceitas livremente, mediante documento firmado entre o <b>PROFESSOR</b> convidado a ministr�-los e a MANTENEDORA;
<br>e) do comparecimento a reuni�es did�tico-pedag�gicas, de avalia��o e de planejamento, quando realizadas fora de seu hor�rio habitual de trabalho, desde que aceito livremente pelo PROFESSOR.
</blockquote>
<b>Par�grafo terceiro</b> � A participa��o em Comiss�es Internas e Externas da Unidade de Ensino da MANTENEDORA, desde que aceita livremente pelo <b>PROFESSOR</b> mediante documento firmado entre a <b>MANTENEDORA</b> e o PROFESSOR, ser� remunerada como aula ou hora normal, acrescida de DSR.

<tr><td class=titulo>11. JANELAS
<tr><td class=campo style="text-align:justify">Considera-se janela a aula vaga existente no hor�rio do <b>PROFESSOR</b> entre duas outras aulas ministradas no mesmo turno. O pagamento da janela � obrigat�rio, devendo o <b>PROFESSOR</b> permanecer � disposi��o da <b>MANTENEDORA</b> neste per�odo, ressalvada a aceita��o pelo PROFESSOR, atrav�s de acordo formalizado entre as partes antes do in�cio das aulas, quando as janelas n�o ser�o pagas.
<br><b>Par�grafo �nico</b> � Ocorrendo a hip�tese da ressalva supra e caso o <b>PROFESSOR</b> seja solicitado esporadicamente a ministrar aulas ou a desenvolver qualquer outra atividade inerente ao magist�rio, no hor�rio de janelas n�o-pagas, essas atividades ser�o remuneradas como aulas extras, com adicional de 100% (cem por cento).

<tr><td class=titulo>12. ADICIONAL POR ATIVIDADES EM OUTROS MUNIC�PIOS
<tr><td class=campo style="text-align:justify">Quando o <b>PROFESSOR</b> desenvolver suas atividades a servi�o da mesma <b>MANTENEDORA</b> em munic�pio diferente daquele onde foi contratado e onde ocorre a presta��o habitual do trabalho, dever� receber um adicional de 25% (vinte e cinco por cento) sobre o total de sua remunera��o no novo munic�pio. Quando o <b>PROFESSOR</b> voltar a prestar servi�os no munic�pio de origem, cessar� a obriga��o no pagamento do adicional.
<br><b>Par�grafo primeiro</b> � Nos casos em que ocorrer a transfer�ncia definitiva do PROFESSOR, aceita livremente por este, em documento firmado entre as partes, n�o haver� a incid�ncia do adicional referido no caput, obrigando-se a <b>MANTENEDORA</b> a efetuar o pagamento de um �nico sal�rio mensal integral, ao PROFESSOR, no ato da transfer�ncia, a t�tulo de ajuda de custo.
<br><b>Par�grafo segundo</b> � Fica assegurada a garantia de emprego pelo per�odo de seis meses ao <b>PROFESSOR</b> transferido de munic�pio, contados a partir do in�cio do trabalho e/ou da efetiva��o da transfer�ncia.
<br><b>Par�grafo terceiro</b> � Caso a <b>MANTENEDORA</b> desenvolva atividade acad�mica em munic�pios considerados conurbados, poder� solicitar isen��o do pagamento do adicional determinado no caput, desde que encaminhe material comprobat�rio ao SEMESP e SEMESP � S�O JOS� DO RIO PRETO, para an�lise e delibera��o do Foro Conciliat�rio para Solu��o de Conflitos Coletivos, previsto na cl�usula 46 da presente Conven��o.

<tr><td class=titulo>13. COMPOSI��O DO SAL�RIO MENSAL DO PROFESSOR
<tr><td class=campo style="text-align:justify">O sal�rio do <b>PROFESSOR</b> � composto, no m�nimo, por tr�s itens: o sal�rio base, o descanso semanal remunerado (DSR)e a hora-atividade.
<br>O sal�rio base � calculado pela seguinte equa��o: n�mero de aulas semanais multiplicado por 4,5 semanas e multiplicado, ainda, pelo valor da hora-aula (artigo 320, par�grafo 1� da CLT).
<br>O DSR corresponde a 1/6 (um sexto) do sal�rio base, acrescido, quando houver, do total de horas extras e do adicional noturno (Lei 605/49).
<br>A hora-ativida decorresponde a 5% (cinco por cento) do total obtido com a somat�ria de todos os valores acima referidos.
<br><b>Par�grafo �nico</b> � A remunera��o adicional do <b>PROFESSOR</b> pelo exerc�cio concomitante de fun��o n�o-docente obedecer� aos crit�rios estabelecidos entre a <b>MANTENEDORA</b> e o <b>PROFESSOR</b> que aceitar o cargo.

<tr><td class=titulo>14. Dura��o da Hora-aula
<tr><td class=campo style="text-align:justify">A dura��o da hora-aula poder� ser de, no m�ximo, cinq�enta minutos.
<br><b>Par�grafo primeiro</b> � Como exce��o ao disposto no caput, a hora-aula poder� ter a dura��o de sessenta minutos nos cursos tecnol�gicos, desde que tenham sido autorizados ou reconhecidos com essa determina��o expressa e cujos PROFESSORES desses cursos tenham sido contratados nessa condi��o.
<br><b>Par�grafo segundo</b> � As MANTENEDORAS de Institui��es de Ensino que possuam cursos tecnol�gicos nas condi��es definidas no par�grafo 1� desta cl�usula dever�o apresentar � Comiss�o Permanente de Negocia��o definida na cl�usula 47 da presente Conven��o, at� o dia 15 de agosto de 2005, a documenta��o de autoriza��o ou reconhecimento do curso com a determina��o expressa de hora-aula com dura��o de sessenta minutos sob pena de, em n�o o fazendo, estar sujeita � majora��o do valor do sal�rio-aula de acordo com o que estabelece o par�grafo 4� desta cl�usula.
<br><b>Par�grafo terceiro</b> � Caso a Comiss�o Permanente de Acompanhamento delibere n�o ter havido determina��o expressa do Minist�rio da Educa��o para que a dura��o da hora-aula dos cursos tecnol�gicos seja de sessenta minutos, a <b>MANTENEDORA</b> dever� majorar o sal�rio-aula de acordo com o que estabelece o par�grafo 4� desta cl�usula.
<br><b>Par�grafo quarto</b> � Em caso de amplia��o da dura��o da hora-aula vigente, respeitado o limite previsto no caput desta cl�usula, a <b>MANTENEDORA</b> dever� acrescer ao sal�rio-aula j� pago, valor proporcional ao acr�scimo do trabalho.

<tr><td class=titulo>15. CARGA HOR�RIA
<tr><td class=campo style="text-align:justify">Quando a <b>MANTENEDORA</b> e o <b>PROFESSOR</b> contratarem carga di�ria de aulas superior aos limites previstos no artigo 318 da CLT, o excedente � carga hor�ria legal ser� remunerado como aula normal, acrescido de DSR, hora-atividade e vantagens pessoais.
<br><b>Par�grafo �nico</b> - Poder� ser flexibilizada a carga hor�ria do <b>PROFESSOR</b> entre jornadas, no exerc�cio de sua fun��o docente e concomitantemente com a atividade administrativa, n�o havendo assim pagamento, no intervalo, de horas aulas e sal�rios, quando o <b>PROFESSOR</b> n�o tenha trabalhado no referido intervalo.

<tr><td class=titulo>16. PRAZO PARA PAGAMENTO DE SAL�RIOS
<tr><td class=campo style="text-align:justify">Os sal�rios dever�o ser pagos, no m�ximo, at� o quinto dia �til do m�s subseq�ente ao trabalhado.
<br><b>Par�grafo �nico</b> � O n�o-pagamento dos sal�rios no prazo obriga a <b>MANTENEDORA</b> a pagar multa di�ria, em favor do PROFESSOR, no valor de 1/50 (um cinq�enta avos) de seu sal�rio mensal.

<tr><td class=titulo>17. DESCONTO DE FALTAS
<tr><td class=campo style="text-align:justify">Na ocorr�ncia de faltas, a <b>MANTENEDORA</b> poder� descontar do sal�rio do PROFESSOR, no m�ximo, o n�mero de aulas em que o mesmo esteve ausente, o DSR (1/6), a hora-atividade e demais vantagens pessoais proporcionais a estas aulas.
<br><b>Par�grafo �nico</b> � � da compet�ncia e de integral responsabilidade da <b>MANTENEDORA</b> estabelecer mecanismos de controle de faltas e de pontualidade dos PROFESSORES, conforme a legisla��o vigente.

<tr><td class=titulo>18. ATESTADOS M�DICOS E ABONO DE FALTAS
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> est� obrigada a aceitar atestados fornecidos por m�dicos ou dentistas credenciados pelo SINPRO, SUS ou, ainda, profissionais conveniados com a pr�pria MANTENEDORA.
<br><b>Par�grafo �nico</b> � Tamb�m ser�o aceitos atestados que tenham sido convalidados pelos profissionais de sa�de do departamento m�dico ou odontol�gico do SINPRO ou conveniados a ele.

<tr><td class=titulo>19. ANOTA��ES NA CARTEIRA DE TRABALHO
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> est� obrigada a promover, em quarenta e oito horas, as anota��es nas Carteiras de Trabalho de seus PROFESSORES, ressalvados eventuais prazos mais amplos permitidos por lei.
<br><b>Par�grafo �nico</b> � � obrigat�ria a anota��o na Carteira de Trabalho das mudan�as provocadas por ascens�o em plano de carreira ou altera��o de titula��o.

<tr><td class=titulo>20. MUDAN�A DE DISCIPLINA
<tr><td class=campo style="text-align:justify">O <b>PROFESSOR</b> n�o poder� ser transferido de uma disciplina para outra, salvo com seu consentimento expresso e por escrito, sob pena de nulidade da referida transfer�ncia.

<tr><td class=titulo>21. Redu��o de carga hor�ria por extin��o ou supress�o de disciplina, classe ou turma
<tr><td class=campo style="text-align:justify">Ocorrendo supress�o de disciplina, classe ou turma, em virtude de altera��o na estrutura curricular prevista ou autorizada pela legisla��o vigente ou por dispositivo regimental devidamente aprovado por �rg�o colegiado da Institui��o de Ensino, o <b>PROFESSOR</b> da disciplina classe ou turma dever� ser comunicado da redu��o da sua carga hor�ria, por escrito, com anteced�ncia m�nima de 30 (trinta) dias do in�cio do per�odo letivo e ter� prioridade para preenchimento de vaga existente em outra classe ou turma ou em outra disciplina para a qual possua habilita��o legal.
<br><b>Par�grafo primeiro</b> � O <b>PROFESSOR</b> dever� manifestar por escrito, no prazo m�ximo de 5 (cinco) dias ap�s a comunica��o da MANTENEDORA, a n�o-aceita��o da transfer�ncia de disciplina ou de classe ou turma ou da redu��o parcial de sua carga hor�ria. A aus�ncia de manifesta��o do <b>PROFESSOR</b> caracterizar� a sua aceita��o.
<br><b>Par�grafo segundo</b> � Caso o <b>PROFESSOR</b> n�o aceite a transfer�ncia para outra disciplina, classe ou turma ou a redu��o parcial de carga hor�ria, a <b>MANTENEDORA</b> dever� manter a carga hor�ria semanal existente ou, em caso contr�rio, proceder � rescis�o do contrato de trabalho, por demiss�o sem justa causa.

<tr><td class=titulo>22. REDU��O DE CARGA HOR�RIA POR DIMINUI��O DE CARGA HOR�RIA POR DIMINUI��O DO N�MERO DE ALUNOS MATRICULADOS 
<tr><td class=campo style="text-align:justify">Na ocorr�ncia de diminui��o do n�mero de alunos matriculados que venha a caracterizar a supress�o de turmas, curso ou disciplina, o <b>PROFESSOR</b> do curso em quest�o dever� ser comunicado, por escrito, da redu��o parcial ou total de sua carga hor�ria at� o final da segunda semana de aulas do per�odo letivo.
<br><b>Par�grafo primeiro</b> - O <b>PROFESSOR</b> dever� manifestar, tamb�m por escrito, a aceita��o ou n�o da redu��o parcial de carga hor�ria no prazo m�ximo de cinco dias ap�s a comunica��o da MANTENEDORA. A aus�ncia de manifesta��o do <b>PROFESSOR</b> caracterizar� a sua n�o-aceita��o.
<br><b>Par�grafo segundo</b> - Caso o <b>PROFESSOR</b> aceite a redu��o parcial de carga hor�ria, dever� formalizar documento junto � <b>MANTENEDORA</b> e, em n�o aceitando, a <b>MANTENEDORA</b> dever� proceder � rescis�o do contrato de trabalho, por demiss�o sem justa causa, caso seja mantida a redu��o parcial de carga hor�ria.
<br><b>Par�grafo terceiro</b> - Na hip�tese de rescis�o contratual, por demiss�o sem justa causa, o aviso pr�vio ser� indenizado, estando a <b>MANTENEDORA</b> desobrigada do pagamento do disposto na cl�usula 29 da presente Conven��o - Garantia Semestral deSal�rios.
<br><b>Par�grafo quarto</b> - N�o ocorrendo redu��o do n�mero de alunos matriculados que venha a caracterizar supress�o do curso, de turma ou de disciplina, a <b>MANTENEDORA</b> que reduzir a carga hor�ria do <b>PROFESSOR</b> estar� sujeita ao disposto na cl�usula 29 - Garantia Semestral de Sal�rios - quando ocorrer a rescis�o do contrato de trabalho do PROFESSOR.

<tr><td class=titulo>23. ABONO DE FALTAS POR CASAMENTO OU LUTO
<tr><td class=campo style="text-align:justify">N�o ser�o descontadas, no curso de nove dias corridos, as faltas do PROFESSOR, por motivo de gala ou luto, este em decorr�ncia de falecimento de pai, m�e, filho, c�njuge, companheira (o) e dependente juridicamente reconhecido.

<tr><td class=titulo>24. IRREDUTIBILIDADE SALARIAL
<tr><td class=campo style="text-align:justify">� proibida a redu��o de remunera��o mensal ou de carga hor�ria, ressalvada a ocorr�ncia do disposto nas cl�usulas 21 e 22 da presente Conven��o, ou ainda, quando ocorrer iniciativa expressa do PROFESSOR. Em qualquer hip�tese, � obrigat�ria a concord�ncia rec�proca, firmada por escrito.
<br><b>Par�grafo primeiro</b> - N�o havendo concord�ncia rec�proca, a parte que deu origem � redu��o prevista nesta cl�usula arcar� com a responsabilidade da rescis�o contratual.
<br><b>Par�grafo segundo</b> � Outras atividades, ainda que inerentes ao trabalho docente, que n�o sejam as de ministrar aulas, de dura��o tempor�ria e determinada, poder�o ser regulamentadas por contrato entre as partes, contendo a caracteriza��o da atividade, o in�cio e a previs�o do t�rmino. 

<tr><td class=titulo>25. UNIFORMES
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> dever� fornecer gratuitamente dois uniformes por ano, quando o seu uso for exigido.

<tr><td class=titulo>26. LICEN�A SEM REMUNERA��O
<tr><td class=campo style="text-align:justify">O <b>PROFESSOR</b> com mais de cinco anos ininterruptos de servi�o na <b>MANTENEDORA</b> ter� direito a licenciar-se, sem direito � remunera��o, por um per�odo m�ximo de dois anos, n�o sendo este per�odo de afastamento computado para contagem de tempo de servi�o ou para qualquer outro efeito, inclusive legal.
<br><b>Par�grafo primeiro</b> � A licen�a ou sua prorroga��o dever� ser comunicada por escrito, � MANTENEDORA, com anteced�ncia m�nima de noventa dias do per�odo letivo, devendo especificar as datas de in�cio e t�rmino do afastamento. A licen�a s� ter� in�cio a partir da data expressa no comunicado, mantendo-se, at� a�, todas as vantagens contratuais. A inten��o de retorno do <b>PROFESSOR</b> � atividade dever� ser comunicada � MANTENEDORA, no m�nimo, sessenta dias antes do t�rmino do afastamento.
<br><b>Par�grafo segundo</b> � O t�rmino do afastamento dever� coincidir com o in�cio do per�odo letivo.
<br><b>Par�grafo terceiro</b> � O <b>PROFESSOR</b> que tenha ou exer�a cargo de confian�a dever�, junto com o comunicado de licen�a, solicitar seu desligamento do cargo a partir do in�cio do per�odo de licen�a.
<br><b>Par�grafo quarto</b> � Considera-se demission�rio o <b>PROFESSOR</b> que, ao t�rmino do afastamento, n�o retornar �s atividades docentes.
<br><b>Par�grafo quinto</b> � Ocorrendo a dispensa sem justa causa ao t�rmino da licen�a, o <b>PROFESSOR</b> n�o ter� direito � Garantia Semestral de Sal�rios, prevista na cl�usula 29 da presente Conven��o.

<tr><td class=titulo>27. LICEN�A � PROFESSORA ADOTANTE
<tr><td class=campo style="text-align:justify">Nos termos da Lei 10.421, de 15 de abril de 2002, ser� assegurada licen�a maternidade � professora que vier a adotar ou obtiver guarda judicial de crian�as, garantido o emprego no per�odo em que a licen�a for concedida.

<tr><td class=titulo>28. LICEN�A PATERNIDADE
<tr><td class=campo style="text-align:justify">A licen�a paternidade ter� dura��o de cinco dias.

<tr><td class=titulo>29. GARANTIA SEMESTRAL DE SAL�RIOS
<tr><td class=campo style="text-align:justify">Ao <b>PROFESSOR</b> demitido sem justa causa, a <b>MANTENEDORA</b> garantir�:
<blockquote style="margin-top:0;margin-bottom:0">
�no primeiro semestre, a partir de 1� de janeiro, os sal�rios integrais at� o dia 30 de junho; 
<br>�no segundo semestre, os sal�rios integrais at� o dia 31 de dezembro, ressalvado o par�grafo 4�. 
</blockquote>
<b>Par�grafo primeiro</b> - N�o ter� direito � Garantia Semestral de Sal�rios o <b>PROFESSOR</b> que, na data da comunica��o da dispensa, contar com menos de 18 (dezoito) meses de servi�o prestado � MANTENEDORA, ressalvado o par�grafo 4� desta cl�usula.
<br><b>Par�grafo segundo</b> � No caso de demiss�es efetuadas no final do primeiro semestre letivo, para n�o ficar obrigada a pagar ao <b>PROFESSOR</b> os sal�rios do segundo semestre a <b>MANTENEDORA</b> dever� observar as seguintes disposi��es:
<blockquote style="margin-top:0;margin-bottom:0">
�com aviso pr�vio a ser trabalhado, a demiss�o dever� ser formalizada com anteced�ncia m�nima de trinta dias do in�cio das f�rias; 
<br>�sendo o aviso pr�vio indenizado, a demiss�o dever� ser formalizada at� um dia antes do in�cio das f�rias, ainda que as f�rias tenham seu in�cio programado para o m�s de julho, obedecendo ao que disp�e a cl�usula 38 da presente Conven��o. 
<br>�Os dias de aviso pr�vio que forem indenizados n�o contar�o como tempo de servi�o para efeito do pagamento da Garantia Semestral de Sal�rios, conforme o estabelecido nesta cl�usula.
</blockquote>
<b>Par�grafo terceiro</b> - No caso de demiss�es efetuadas no final do ano letivo, para n�o ficar obrigada a pagar ao <b>PROFESSOR</b> os sal�rios do primeiro semestre do ano seguinte a <b>MANTENEDORA</b> dever� observar as seguintes disposi��es:
<blockquote style="margin-top:0;margin-bottom:0">
�com aviso pr�vio a ser trabalhado, a demiss�o dever� ser formalizada com anteced�ncia m�nima de trinta dias do in�cio do recesso escolar; 
<br>�sendo o aviso pr�vio indenizado, a demiss�o dever� ser formalizada at� um dia antes do in�cio do recesso escolar. 
<br>�Os dias de aviso pr�vio que forem indenizados n�o contar�o como tempo de servi�o para efeito do pagamento da Garantia Semestral de Sal�rios, conforme o estabelecido nesta cl�usula.
</blockquote>
<b>Par�grafo quarto</b> - Quando as demiss�es ocorrerem a partir de 16 de outubro, a <b>MANTENEDORA</b> pagar�, independentemente do tempo de servi�o do professor, valor correspondente � remunera��o devida at� o dia 18 de janeiro do ano subseq�ente, inclusive, ressalvados os contratos de experi�ncia e por prazo determinado, estes �ltimos v�lidos somente nos casos de substitui��o tempor�ria, conforme o disposto na al�nea a) do par�grafo 2� da cl�usula 10� da presente Conven��o.
<br><b>Par�grafo quinto</b> - Os sal�rios complementares previstos nesta cl�usula ter�o natureza indenizat�ria, n�o integrando, para nenhum efeito legal, o tempo de servi�o do professor.
<br><b>Par�grafo sexto</b> - O aviso pr�vio de trinta dias previsto no artigo 487 da CLT j� est� integrado �s indeniza��es tratadas nesta cl�usula.

<tr><td class=titulo>30. GARANTIA DE EMPREGO � GESTANTE
<tr><td class=campo style="text-align:justify">� proibida a dispensa arbitr�ria ou sem justa causa da PROFESSORA gestante, desde o in�cio da gravidez at� sessenta dias ap�s o t�rmino do afastamento legal.
O aviso pr�vio come�ar� a contar a partir do t�rmino do per�odo de estabilidade.

<tr><td class=titulo>31. CRECHES
<tr><td class=campo style="text-align:justify">� obrigat�ria a instala��o de local destinado � guarda de crian�as de at� seis meses, quando a <b>MANTENEDORA</b> mantiver contratadas, em jornada integral, pelo menos trinta funcion�rias com idade superior a 16 anos. A manuten��o da creche poder� ser substitu�da pelo pagamento do reembolso-creche, nos termos da legisla��o em vigor (artigo 389, par�grafo 1� da CLT e Portarias MTb n� 3296 de 03.09.86 e n� 670, de 27/08/97), ou ainda, a celebra��o de conv�nio com uma entidade reconhecidamente id�nea.

<tr><td class=titulo>32. GARANTIAS AO <b>PROFESSOR</b> EM VIAS DE APOSENTADORIA
<tr><td class=campo style="text-align:justify">Fica assegurado ao <b>PROFESSOR</b> que, comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria integral por tempo de servi�o ou da aposentadoria por idade, a garantia de emprego durante o per�odo que faltar at� a aquisi��o do direito.
<br><b>Par�grafo primeiro</b> � A garantia de emprego � devida ao <b>PROFESSOR</b> que esteja contratado pela <b>MANTENEDORA</b> h� pelo menos tr�s anos.
<br><b>Par�grafo segundo</b> � A comprova��o � <b>MANTENEDORA</b> dever� ser feita mediante a apresenta��o de documento que ateste o tempo de servi�o. Este documento dever� ser emitido pela Previd�ncia Social ou por pessoa credenciada junto ao �rg�o previdenci�rio. Se o <b>PROFESSOR</b> depender de documenta��o para realiza��o da contagem, ter� um prazo de quarenta e cinco dias, a contar da data da comunica��o da dispensa. Comprovada a solicita��o de tal documenta��o, os prazos ser�o prorrogados at� que a mesma seja emitida.
<br><b>Par�grafo terceiro</b> � O contrato de trabalho do <b>PROFESSOR</b> s� poder� ser rescindido por m�tuo acordo homologado pelo SINPRO ou pedido de demiss�o.
<br><b>Par�grafo quarto</b> � Havendo acordo formal entre as partes, o <b>PROFESSOR</b> poder� exercer outra fun��o, inerente ao magist�rio, durante o per�odo em que estiver garantido pela estabilidade.
<br><b>Par�grafo quinto</b> � O aviso pr�vio, em caso de demiss�o sem justa causa, integra o per�odo de estabilidade previsto nesta cl�usula.

<tr><td class=titulo>33. MULTA POR ATRASO NA HOMOLOGA��O
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> deve homologar a rescis�o contratual no dia seguinte ao t�rmino do aviso pr�vio, quando trabalhado, ou dez dias ap�s o desligamento, quando houver dispensa do cumprimento de aviso pr�vio. O atraso na homologa��o obrigar� a <b>MANTENEDORA</b> ao pagamento de multa, em favor do PROFESSOR, correspondente a um m�s de sua remunera��o, conforme o disposto no par�grafo 8� do artigo 477 da CLT. A partir do vig�simo dia de atraso, haver� ainda multa di�ria de 0,2% (dois d�cimos percentuais) do sal�rio mensal. A <b>MANTENEDORA</b> est� desobrigada de pagar a multa quando o atraso vier a ocorrer, comprovadamente, por motivos alheios a sua vontade.
<br><b>Par�grafo �nico</b> � O SINPRO est� obrigado a fornecer comprovante de comparecimento sempre que a <b>MANTENEDORA</b> se apresentar para homologa��o das rescis�es contratuais e comprovar a convoca��o do PROFESSOR.

<tr><td class=titulo>34. DEMISS�O POR JUSTA CAUSA
<tr><td class=campo style="text-align:justify">Quando houver demiss�o por justa causa, nos termos do art. 482 da CLT, a <b>MANTENEDORA</b> est� obrigada a determinar na carta-aviso o motivo que deu origem � dispensa. Caso contr�rio, fica descaracterizada a justa causa.

<tr><td class=titulo>35. READMISS�O DO PROFESSOR
<tr><td class=campo style="text-align:justify">O <b>PROFESSOR</b> que for readmitido at� doze meses ap�s o seu desligamento ficar� desobrigado de firmar contrato de experi�ncia.

<tr><td class=titulo>36. INDENIZA��ES POR DISPENSA IMOTIVADA
<tr><td class=campo style="text-align:justify">O <b>PROFESSOR</b> demitido sem justa causa ter� direito a uma indeniza��o, al�m do aviso pr�vio legal de trinta dias e das indeniza��es previstas na cl�usula 28 desta Conven��o, quando forem devidas, nas condi��es abaixo especificadas:
<blockquote style="margin-top:0;margin-bottom:0">
a) tr�s (03) dias para cada ano trabalhado na MANTENEDORA;
<br>b) aviso pr�vio adicional de quinze dias, caso o <b>PROFESSOR</b> tenha, no m�nimo, cinq�enta anos de idade e que, � data do desligamento, conte com pelo menos um ano de servi�o na MANTENEDORA.
</blockquote>
<b>Par�grafo primeiro</b> � N�o estar� obrigada ao pagamento da indeniza��o prevista na al�nea a) a <b>MANTENEDORA</b> que tiver garantido ao <b>PROFESSOR</b> demitido, durante pelo menos um ano, pagamento mensal de adicional por tempo de servi�o decorrente de plano de cargos e sal�rios ou de anu�nio, q�inq��nio ou equivalente, cujo valor corresponda a, no m�nimo, 1% (um por cento) do valor da hora-aula por ano trabalhado e, por conseq��ncia, do sal�rio mensal.
<br><b>Par�grafo segundo</b> � N�o ter� direito � indeniza��o assegurada na al�nea b) do caput, o <b>PROFESSOR</b> que, na data de admiss�o na MANTENEDORA, contar com mais de cinq�enta anos de idade.
<br><b>Par�grafo terceiro</b> � Para fazer jus � isen��o prevista no par�grafo primeiro desta cl�usula, a <b>MANTENEDORA</b> dever� encaminhar � Comiss�o Permanente de Negocia��o definida na cl�usula 46 desta Conven��o, no prazo m�ximo de noventa dias a contar da data da assinatura da presente Conven��o, documenta��o que comprove o plano de pagamento de adicional por tempo de servi�o nas condi��es estabelecidas no referido par�grafo.
<br><b>Par�grafo quarto</b> � Para a <b>MANTENEDORA</b> que n�o estiver enquadrada nos par�grafos primeiro e segundo, o pagamento das verbas indenizat�rias previstas nesta cl�usula n�o ser� cumulativo, cabendo ao PROFESSOR, no desligamento, o maior valor monet�rio entre os previstos nas al�neas a) e b) do caput.
<br><b>Par�grafo quinto</b> � Essas indeniza��es n�o contar�o, para nenhum efeito, como tempo de servi�o.

<tr><td class=titulo>37. ATESTADOS DE AFASTAMENTO E SAL�RIOS
<tr><td class=campo style="text-align:justify">Sempre que solicitada, a <b>MANTENEDORA</b> dever� fornecer ao <b>PROFESSOR</b> atestado de afastamento e sal�rio (AAS), previsto na legisla��o previdenci�ria.

<tr><td class=titulo>38. F�RIAS
<tr><td class=campo style="text-align:justify">As f�rias anuais dos PROFESSORES ser�o coletivas, com dura��o de trinta dias corridos e gozadas em julho de 2005 e julho de 2006. Qualquer altera��o dever� ser aprovada por �rg�o competente, conforme o estabelecido em Estatuto ou Regimento e dever� constar do calend�rio escolar.
<br><b>Par�grafo primeiro</b> � A <b>MANTENEDORA</b> est� obrigada a pagar o sal�rio das f�rias e o abono constitucional de 1/3 (um ter�o) at� quarenta e oito horas antes do in�cio das f�rias.
<br><b>Par�grafo segundo</b> � As f�rias n�o poder�o ser iniciadas aos domingos, feriados, dias de compensa��o do descanso semanal remunerado e nem aos s�bados, quando estes n�o forem dias normais de aula.

<tr><td class=titulo>39. RECESSO ESCOLAR
<tr><td class=campo style="text-align:justify">O recesso escolar anual � obrigat�rio e tem dura��o de trinta dias corridos. Na vig�ncia da presente Conven��o os recessos escolares ser�o gozados preferencialmente no m�s de janeiro de 2006 e no m�s de janeiro de 2007.
Durante o recesso escolar anual que n�o pode, de maneira alguma, coincidir com o per�odo definido para as f�rias coletivas do ano respectivo, o <b>PROFESSOR</b> n�o poder� ser convocado para nenhum trabalho.
<br><b>Par�grafo primeiro</b> �Na vig�ncia da presente Conven��o, as institui��es cujos calend�rios escolares, determinados pelo �rg�o competente conforme o estabelecido em Estatuto ou Regimento, n�o observarem o determinado pelo caput para o recesso escolar anual dos PROFESSORES, poder�o conced�-lo em um per�odo de, no m�nimo vinte dias corridos e em, no m�ximo, mais dois per�odos com igual n�mero de dias corridos, desde que observem as seguintes condi��es:
<blockquote style="margin-top:0;margin-bottom:0">
a) Vinte dias corridos em janeiro de 2006 e os dois per�odos com igual n�mero de dias corridos, obrigatoriamente no per�odo compreendido entre mar�o de 2005 e fevereiro de 2006.
<br>b) Vinte dias corridos em janeiro de 2007 e os dois per�odos com igual n�mero de dias corridos, obrigatoriamente no per�odo compreendido entre mar�o de 2006 e fevereiro de 2007.
</blockquote>
<b>Par�grafo segundo</b> � No caso dos calend�rios escolares preverem a divis�o do recesso escolar dos PROFESSORES, os per�odos definidos na conformidade do par�grafo primeiro n�o poder�o ser iniciados aos domingos, feriados, dias de compensa��o do descanso semanal remunerado e nem aos s�bados, quando estes n�o forem dias normais de aulas.
<br><b>Par�grafo terceiro</b> � As Institui��es cujas atividades n�o possam ser interrompidas, tais como aquelas desenvolvidas em hospital, cl�nica, laborat�rio de an�lise, escrit�rios experimentais, pesquisas, dentre outros, ou que ministrem cursos em que sejam utilizadas instala��es espec�ficas ou que prestem atendimento � comunidade que n�o pode ser suspenso, poder�o conceder aos PROFESSORES o recesso escolar anual definido no caput de maneira escalonada ao longo de cada ano.
<br><b>Par�grafo quarto</b> � Os calend�rios escolares que definir�o os per�odos de recesso escolar dos PROFESSORES ser�o obrigatoriamente divulgados aos PROFESSORES at� o in�cio de cada per�odo letivo.

<tr><td class=titulo>40. DELEGADO REPRESENTANTE
<tr><td class=campo style="text-align:justify">Em cada unidade de ensino que tiver mais de cinq�enta PROFESSORES, a <b>MANTENEDORA</b> assegurar� elei��o de um Delegado Representante, que ter� garantia de emprego e sal�rios a partir da inscri��o de sua candidatura at� o t�rmino do semestre letivo em que sua gest�o se encerrar.
<br><b>Par�grafo primeiro</b> � O mandato do Delegado Representanteser� de um ano.
<br><b>Par�grafo segundo</b> � A elei��o do Delegado Representanteser� realizada pelo SINPRO na unidade de ensino da MANTENEDORA, por voto direto e secreto. � exigido quorum de 50% (cinq�enta por cento) mais um do corpo docente da unidade onde a elei��o ocorrer.
<br><b>Par�grafo terceiro</b> � O SINPRO comunicar� a elei��o � <b>MANTENEDORA</b> com anteced�ncia m�nima de sete dias corridos. Nenhum candidato poder� ser demitido a partir da data da comunica��o at� o t�rmino da apura��o.
<br><b>Par�grafo quarto</b> � � condi��o necess�ria que os candidatos tenham, � data da elei��o, pelo menos um ano de servi�o na MANTENEDORA.

<tr><td class=titulo>41. QUADRO DE AVISOS
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> dever� colocar, nas salas de professores, quadro de aviso � disposi��o do SINPRO para fixa��o de comunicados de interesse da categoria, sendo vedada a divulga��o de mat�ria pol�tico-partid�ria ou ofensiva a quem quer que seja.

<tr><td class=titulo>42. ASSEMBL�IAS SINDICAIS
<tr><td class=campo style="text-align:justify">Todo <b>PROFESSOR</b> ter� direito a abono de faltas para o comparecimento a assembl�ias da categoria.
<br><b>Par�grafo primeiro</b> � Na vig�ncia desta Conven��o, os abonos est�o limitados a dois s�bados e mais dois dias �teis para cada per�odo compreendido entre o m�s de mar�o e o m�s de fevereiro do ano subseq�ente. As duas assembl�ias realizadas durante os dias �teis dever�o ocorrer em per�odos distintos.
<br><b>Par�grafo segundo</b> � O SINPRO ou a FEPESP dever� informar ao SEMESP ou � MANTENEDORA, por escrito, com anteced�ncia m�nima de quinze dias corridos. Na comunica��o dever�o constar a data e o hor�rio da assembl�ia.
<br><b>Par�grafo terceiro</b> � Os dirigentes sindicais n�o est�o sujeitos ao limite previsto no par�grafo 1� desta cl�usula. As aus�ncias decorrentes do comparecimento �s assembl�ias de suas entidades ser�o abonadas mediante pr�via comunica��o formal � MANTENEDORA.
<br><b>Par�grafo quarto</b> � A <b>MANTENEDORA</b> poder� exigir dos PROFESSORES e dos dirigentes sindicais atestado emitido pelo SINPRO ou pela FEPESP que comprove
o seu comparecimento � assembl�ia.

<tr><td class=titulo>43. CONGRESSOS, SIMP�SIOS E EQUIVALENTES
<tr><td class=campo style="text-align:justify">Os abonos de falta para comparecimento a congressos e simp�sios ser�o concedidos mediante aceita��o por parte da MANTENEDORA, que dever� formalizar por escrito a dispensa do PROFESSOR.
<br><b>Par�grafo �nico</b> � A participa��o do <b>PROFESSOR</b> nos eventos descritos no caput n�o caracterizar� atividade extraordin�ria.

<tr><td class=titulo>44. CONGRESSO DO SINPRO
<tr><td class=campo style="text-align:justify">Em cada ano da vig�ncia desta Conven��o, o SINPRO promover� um evento de natureza pol�tica ou pedag�gica (congresso ou jornada). A <b>MANTENEDORA</b> abonar� as aus�ncias de seus PROFESSORES que participarem do evento, nos seguintes limites:
<blockquote style="margin-top:0;margin-bottom:0">
a) na unidade de ensino que tenha at� 49 PROFESSORES ser� garantido o abono a um PROFESSOR;
<br>b) na unidade de ensino que tenha entre 50 e 99 PROFESSORES ser� garantido o abono a dois PROFESSORES;
<br>c) na unidade de ensino que tenha mais de cem PROFESSORES ser� garantido o abono a tr�s PROFESSORES.
</blockquote>
Tais faltas, limitadas ao m�ximo em dois dias �teis al�m do s�bado, em cada evento, ser�o abonadas mediante a apresenta��o de atestado de comparecimento fornecido pelo SINPRO. O <b>PROFESSOR</b> dever� repor as aulas que, por ventura, sejam necess�rias para complementa��o das horas letivas m�nimas exigidas pela legisla��o.

<tr><td class=titulo>45. RELA��O NOMINAL
<tr><td class=campo style="text-align:justify">Na vig�ncia desta Conven��o, obriga-se a <b>MANTENEDORA</b> a encaminhar ao SINPRO, at� o final do m�s de junho de cada ano, a rela��o nominal dos PROFESSORES que integram seu quadro de funcion�rios, acompanhada do valor do sal�rio mensal e das guias das contribui��es sindical e assistencial.

<tr><td class=titulo>46. FORO CONCILIAT�RIO PARA SOLU��O DE CONFLITOS COLETIVOS
<tr><td class=campo style="text-align:justify">Fica mantida a exist�ncia do Foro Conciliat�rio que tem como objetivo procurar
resolver quest�es referentes ao n�o cumprimento de normas estabelecidas na presente Conven��o e eventuais diverg�ncias trabalhistas existentes entre a <b>MANTENEDORA</b> e seus PROFESSORES.
<br><b>Par�grafo primeiro</b> � O Foro ser� composto por membros do SEMESP e do SINPRO. As reuni�es dever�o contar, tamb�m, com as partes em conflito que, se assim o desejarem, poder�o delegar representantes para substitu�-las e/ou serem assistidas por advogados.
<br><b>Par�grafo segundo</b> � O SEMESP e o SINPRO dever�o indicar os seus representantes no Foro num prazo de trinta dias a contar da assinatura desta Conven��o.
<br><b>Par�grafo terceiro</b> � Cada se��o do Foro ser� realizada no prazo m�ximo de quinze dias a contar da solicita��o formal e obrigat�ria de qualquer uma das entidades que o comp�em, devendo constar na solicita��o a data, o local e o hor�rio em que a mesma dever� se realizar. O n�o-comparecimento de qualquer uma das partes acarretar� no encerramento imediato das negocia��es.
<br><b>Par�grafo quarto</b> � Nenhuma das partes envolvidas ingressar� com a��o na Justi�a do Trabalho durante as negocia��es de entendimento.
<br><b>Par�grafo quinto</b> � Na aus�ncia de solu��o do conflito ou na hip�tese de n�o comparecimento de qualquer uma das partes, a comiss�o respons�vel pelo Foro fornecer� certid�o atestando o encerramento da negocia��o.
<br><b>Par�grafo sexto</b> � Na hip�tese de sucesso das negocia��es, a crit�rio do Foro, a <b>MANTENEDORA</b> ficar� desobrigada de arcar com a multa prevista na cl�usula 54 desta Conven��o.
<br><b>Par�grafo s�timo</b> � As decis�es do Foro ter�o efic�cia legal entre as partes acordantes. O descumprimento das decis�es assumidas gerar� multa a ser estabelecida no Foro,independentemente daquelas j� estabelecidas nesta Conven��o.
<br><b>Par�grafo oitavo</b> � Na hip�tese de incapacidade econ�mico-financeira das MANTENEDORAS, os casos ser�o remetidos para an�lise e delibera��o deste foro.

<tr><td class=titulo>47. COMISS�O PERMANENTE DE NEGOCIA��O
<tr><td class=campo style="text-align:justify">Fica mantida a Comiss�o Permanente de Negocia��o constitu�da de forma parit�ria, por tr�s representantes das entidades sindicais profissional e econ�mica, com o objetivo de:
<blockquote style="margin-top:0;margin-bottom:0">
�fiscalizar o cumprimento das cl�usulas vigentes; 
<br>�elucidar eventuais diverg�ncias de interpreta��o das cl�usulas desta Conven��o; 
<br>�discutir quest�es n�o-contempladas na presente Conven��o. 
<br>�deliberar no prazo m�ximo de trinta dias a contar da data da solicita��o protocolizada no SEMESP, sobre a isen��o prevista na cl�usula 36 da presente Conven��o; sobre modifica��o de pagamento da assist�ncia m�dico-hospitalar, conforme os par�grafos 1� e 3� da cl�usula 49 da presente Conven��o e sobre o valor da remunera��o da hora-aula, conforme o par�grafo 2� da cl�usula 14 da presente Conven��o. 
<br>�criar subs�dios para a Comiss�o de Tratativas Salariais, atrav�s da elabora��o de documentos, para a defini��o das fun��es/atividades e o regime de trabalho dos PROFESSORES. 
</blockquote>
<b>Par�grafo primeiro</b> - As entidades sindicais componentes da Comiss�o Permanente deNegocia��o indicar�o seus representantes, no prazo m�ximo de trinta dias corridos, a contar da assinatura da presente Conven��o.
<br><b>Par�grafo segundo</b> - A Comiss�o Permanente de Negocia��o dever� reunir-se mensalmente, no d�cimo dia �til, �s 15 horas, alternadamente nas sedes das entidades sindicais que a comp�em. No caso espec�fico do item d) do caput, dever� haver convoca��o espec�fica feita pela entidade sindical patronal.

<tr><td class=titulo>48. ACORDOS INTERNOS � CL�USULAS MAIS FAVOR�VEIS
<tr><td class=campo style="text-align:justify">Ficam assegurados os direitos mais favor�veis decorrentes de acordos internos ou de acordos coletivos de trabalho celebrados entre a <b>MANTENEDORA</b> e o SINPRO.

<tr><td class=titulo>49. ASSIST�NCIA M�DICO-HOSPITALAR
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> est� obrigada a assegurar, a suas expensas, assist�ncia m�dico-hospitalar a todos os seus PROFESSORES, sendo-lhe facultada a escolha por plano de sa�de, seguro-sa�de ou conv�nios com empresas prestadoras de servi�os m�dico-hospitalares. Poder� ainda prestar a referida assist�ncia diretamente, em se tratando de institui��es que disponham de servi�os de sa�de e hospitais pr�prios ou conveniados. Qualquer que seja a op��o, a assist�ncia m�dico-hospitalar deve assegurar as condi��es e os requisitos m�nimos que seguem relacionados:
<blockquote style="margin-top:0;margin-bottom:0">
1. Abrang�ncia - A assist�ncia m�dico-hospitalar deve ser realizada no munic�pio onde funciona o estabelecimento de ensino superior ou onde vive o PROFESSOR, a crit�rio da MANTENEDORA. Em casos de emerg�ncia, dever� haver garantia de atendimento integral em qualquer localidade do Estado de S�o Paulo ou fixa��o em contrato, de formas de reembolso.
<br>2. Coberturas m�nimas 
<blockquote style="margin-top:0;margin-bottom:0">
2.1 Quarto para quatro pacientes, no m�ximo.
<br>2.2 Consultas.
<br>2.3 Prazo de interna��o de 365 dias por ano (comum e UTI/CTI).
<br>2.4 Parto independentemente do estado grav�dicio.
<br>2.5 Mol�stias infecto-contagiosas que exijam interna��o.
<br>2.6 Exames laboratoriais, ambulatoriais e hospitalares.
</blockquote>
3. Car�ncia - N�o haver� car�ncia na presta��o dos servi�os m�dicos laboratoriais.
<br>4. <b>PROFESSOR</b> ingressante - N�o haver� car�ncia para o <b>PROFESSOR</b> ingressante, independentemente da data em que for contratado.
<br>5. Pagamento - Caber� ao <b>PROFESSOR</b> o pagamento de 10% (dez por cento) do valor da Assist�ncia M�dica, limitado tal pagamento a R$ 8,00, respeitado o disposto nos par�grafos 1� e 2�.
</blockquote>
<b>Par�grafo primeiro</b> � Caso a assist�ncia m�dico-hospitalar vigente na Institui��o venha a sofrer reajuste em virtude de poss�veis modifica��es estabelecidas em legisla��o que abranja o segmento, Lei 9.656, de 03 de julho de 1998 e MP 2.097-39, de 26 de abril de 2001, ou que vierem a ser estabelecidas em lei, ou por mudan�a de empresa prestadora de servi�o a pedido dos empregados da Institui��o, ou por quebra unilateral de contrato por parte da atual empresa prestadora de servi�o, a <b>MANTENEDORA</b> continuar� a contribuir com o valor mensal vigente at� a data da modifica��o, devendo o <b>PROFESSOR</b> arcar com o valor excedente, que ser� descontado em folha e consignado no comprovante de pagamento, nos termos do artigo 462 da CLT.
<br><b>Par�grafo segundo</b> - Caso ocorra mudan�a de empresa prestadora de servi�o, por decis�o unilateral da MANTENEDORA, com conseq�ente reajuste no valor vigente, o <b>PROFESSOR</b> estar� isento do pagamento do valor excedente, cabendo � <b>MANTENEDORA</b> prover integralmente a assist�ncia m�dico-hospitalar, sem nenhum �nus para o PROFESSOR.
<br><b>Par�grafo terceiro</b> - Para efeito do disposto no � 1� desta cl�usula, caber� � <b>MANTENEDORA</b> remeter a documenta��o comprobat�ria para an�lise e delibera��o da Comiss�o Permanente de Negocia��o, nos termos da cl�usula 47.
<br><b>Par�grafo quarto</b> - Fica facultado ao <b>PROFESSOR</b> optar pela presta��o de assist�ncia m�dico-hospitalar em uma �nica institui��o de ensino, quando mantiver mais de um v�nculo empregat�cio como PROFESSOR. � necess�rio que o <b>PROFESSOR</b> se manifeste por escrito, com anteced�ncia m�nima de vinte dias, para que a <b>MANTENEDORA</b> possa proceder � suspens�o dos servi�os.
<br><b>Par�grafo quinto</b> � Caso o <b>PROFESSOR</b> mantenha v�nculo empregat�cio com mais de uma Institui��o de Ensino, as MANTENEDORAS, em conjunto, poder�o optar por conceder-lhe um �nico plano de sa�de, pago por elas em regime de cotiza��o de custos, respeitadas as condi��es estabelecidas nesta cl�usula.
<br><b>Par�grafo sexto</b> - Mediante pagamento complementar e ades�o facultativa devidamente documentada, o <b>PROFESSOR</b> poder� optar pela amplia��o dos servi�os de sa�de garantidos nesta Conven��o ou estend�-los a seus dependentes.

<tr><td class=titulo>50. Bolsas de Estudo
<tr><td class=campo style="text-align:justify">Todo <b>PROFESSOR</b> tem direito a bolsas de estudo integrais, incluindo matr�cula, para si, seus filhos ou dependentes legais, estes �ltimos entendidos como aqueles reconhecidos pela legisla��o do Imposto de Renda ou aqueles que estejam sob a guarda judicial do <b>PROFESSOR</b> e vivam sob sua depend�ncia econ�mica, devidamente comprovada. Os filhos do <b>PROFESSOR</b> poder�o usufruir as bolsas de estudo integrais, sem qualquer �nus, desde que n�o tenham 25 (vinte e cinco) anos completos ou mais na data de realiza��o do exame vestibular ou do processo seletivo que define o ingresso no curso superior.
<br>As bolsas de estudo s�o v�lidas para cursos de gradua��o, p�s-gradua��o ou seq�enciais existentes e administrados pela <b>MANTENEDORA</b> para a qual o <b>PROFESSOR</b> trabalha, observado o disposto nesta cl�usula e par�grafos seguintes.
<br><b>Par�grafo primeiro</b> � O direito �s bolsas de estudo passa a vigorar ao t�rmino do contrato de experi�ncia, cuja dura��o n�o pode exceder de 90 (noventa) dias, conforme par�grafo �nico do artigo 445 da CLT.
<br><b>Par�grafo segundo</b> - A <b>MANTENEDORA</b> est� obrigada a conceder duas bolsas de estudo, sendo que, nos cursos de gradua��o ou seq�enciais, n�o ser� poss�vel que o bolsista conclua mais de um curso nesta condi��o.
<br><b>Par�grafo terceiro</b> - A utiliza��o do benef�cio previsto nesta cl�usula � transit�ria e n�o-habitual e, por isso, n�o possui car�ter remunerat�rio e nem se vincula, para nenhum efeito, ao sal�rio ou remunera��o percebida pelo <b>PROFESSOR</b> nos termos do inciso XIX, do par�grafo 9� do artigo 214 do Decreto 3048, de 06 de maio de 1999 e do par�grafo 2� do artigo 458 da CLT, com a reda��o dada pela Lei 10243, de 19 de junho de 2001.
<br><b>Par�grafo quarto</b> - As bolsas de estudo ser�o mantidas quando o <b>PROFESSOR</b> estiver licenciado para tratamento de sa�de ou em gozo de licen�a mediante anu�ncia da MANTENEDORA, excetuado o disposto na cl�usula 26 da presente Conven��o � Licen�a sem Remunera��o.
<br><b>Par�grafo quinto</b> - No caso de falecimento do PROFESSOR, os dependentes que j� se encontram estudando em estabelecimento de ensino superior da <b>MANTENEDORA</b> continuar�o a gozar das bolsas de estudo at� o final do curso, ressalvado o disposto no par�grafo 8� desta cl�usula.
<br><b>Par�grafo sexto</b> - No caso de dispensa sem justa causa durante o per�odo letivo, ficam garantidas ao PROFESSOR, at� o final do per�odo letivo, as bolsas de estudo j� existentes.
<br><b>Par�grafo s�timo</b> - As bolsas de estudo integrais em cursos de p�s-gradua��o ou especializa��o existentes e administrados pela <b>MANTENEDORA</b> s�o v�lidas exclusivamente para o PROFESSOR, em �reas correlatas �s disciplinas que o mesmo ministra na Institui��o ou que visem a capacita��o docente, respeitados os crit�rios de sele��o exigidos para ingresso no mesmo e obedecer�o as seguintes condi��es:
<blockquote style="margin-top:0;margin-bottom:0">
�os cursos stricto sensu ou de especializa��o que fixem um n�mero m�ximo de alunos por turma, s�o limitadas em 30% (trinta por cento) do total de vagas oferecidas; 
<br>�nos cursos de p�s-gradua��o lato sensu n�o haver� limites de vagas. Caso a estrutura do curso torne necess�ria a limita��o do n�mero de alunos ser� observado o disposto na al�nea a) deste par�grafo. 
</blockquote>
<b>Par�grafo oitavo</b> - Os bolsistas que forem reprovados no per�odo letivo perder�o o direito � bolsa de estudo, voltando a gozar do benef�cio quando lograrem aprova��o no referido per�odo. As disciplinas cursadas em regime de depend�ncia ser�o de total responsabilidade do bolsista, arcando o mesmo com o seu custo.
<br><b>Par�grafo nono</b> - Considera-se adquirido o direito daquele <b>PROFESSOR</b> que j� esteja usufruindo bolsas de estudo em n�mero superior ao definido nesta cl�usula.

<tr><td class=titulo>51. AUTORIZA��O PARA DESCONTO EM FOLHA DE PAGAMENTO
<tr><td class=campo style="text-align:justify">O desconto do <b>PROFESSOR</b> em folha de pagamento somente poder� ser realizado mediante sua autoriza��o, nos termos dos artigos 462 e 545 da CLT, quando os valores forem destinados ao custeio de pr�mios de seguro, planos de sa�de, mensalidades associativas ou outras que constem da sua expressa autoriza��o, desde que n�o haja previs�o expressa de desconto na presente norma coletiva.
<br><b>Par�grafo �nico</b> � Encontra-se no SINPRO, � disposi��o da MANTENEDORA, c�pia de autoriza��o do <b>PROFESSOR</b> para o desconto da mensalidade associativa.

<tr><td class=titulo>52. ESTABILIDADE PARA PORTADORES DE DOEN�AS GRAVES
<tr><td class=campo style="text-align:justify">Fica assegurada, at� alta m�dica, considerada como apto ao trabalho, ou eventual concess�o de aposentadoria por invalidez,estabilidade no emprego aos PROFESSORES acometidos por doen�as graves ou incur�veis e aos PROFESSORES portadores do v�rus HIV que vierem a apresentar qualquer tipo de infec��o ou doen�a oportunista, resultante da patologia de base.
<br><b>Par�grafo �nico</b> - S�o consideradas doen�as graves ou incur�veis, a tuberculose ativa, aliena��o mental, esclerose m�ltipla, neoplasia maligna, cegueira definitiva, hansen�ase, cardiopatia grave, doen�a de Parkinson, paralisia irrevers�vel e incapacitante, espondiloartrose anquilosante, nefropatia grave, estados do Mal de Paget (oste�te deformante) e contamina��o grave por radia��o.

<tr><td class=titulo>53. GARANTIAS AO <b>PROFESSOR</b> COM SEQ�ELAS E READAPTA��O
<tr><td class=campo style="text-align:justify">Ser� garantida ao <b>PROFESSOR</b> acidentado no trabalho ou acometido por doen�a profissional a perman�ncia na empresa em fun��o compat�vel com o seu estado f�sico, sem preju�zo na remunera��o antes percebida, desde que, ap�s o acidente ou comprova��o da aquisi��o de doen�a profissional, apresente, cumulativamente, redu��o da capacidade laboral, atestada pelo �rg�o oficial e que se tenha tornado incapaz de exercer a fun��o que anteriormente desempenhava, obrigado, por�m, o <b>PROFESSOR</b> nessa situa��o a participar dos processos de readapta��o e reabilita��o profissional.
<br><b>Par�grafo �nico</b> � O per�odo de estabilidade do <b>PROFESSOR</b> que se encontre participando dos processos de readapta��o e reabilita��o profissional ser� o previsto em lei.

<tr><td class=titulo>54. MULTA POR DESCUMPRIMENTO DA CONVEN��O
<tr><td class=campo style="text-align:justify">O descumprimento desta Conven��o obrigar� a <b>MANTENEDORA</b> ao pagamento de multa correspondente a 1% (um por cento) do sal�rio do PROFESSOR, para cada uma das cl�usulas n�o-cumpridas, acrescidas de juros, a cada <b>PROFESSOR</b> prejudicado.
<br><b>Par�grafo �nico</b> � A <b>MANTENEDORA</b> est� desobrigada de arcar com a multa prevista nesta cl�usula, caso o artigo da Conven��o j� estabele�a uma multa pelo n�o-cumprimento da mesma.

<tr><td class=titulo>55. CONTRIBUI��O ASSISTENCIAL PROFISSIONAL
<tr><td class=campo style="text-align:justify">Obriga-se a <b>MANTENEDORA</b> a promover o desconto nos exerc�cios de 2005 e 2006, na folha de pagamento de seus PROFESSORES, sindicalizados e/ou filiados ou n�o, para recolhimento em favor do SINPRO, entidade legalmente representativa da categoria dos PROFESSORES, na base territorial conferida pela respectiva carta sindical oupelo inciso I, artigo 8� da Constitui��o Federal, em conta especial, da import�ncia correspondente ao percentual estabelecido ou ao que vier a ser estabelecido na Assembl�ia Geral da categoria. O recolhimento ser� realizado obrigatoriamente pela pr�pria MANTENEDORA, em guias pr�prias, acompanhadas das correspondentes rela��es nominais e valores devidos. As import�ncias destinam-se � cria��o, manuten��o e amplia��o dos servi�os assistenciais do SINPRO, na conformidade das assembl�ias gerais.
<br><b>Par�grafo primeiro</b> - Quando a <b>MANTENEDORA</b> deixar de efetuar o recolhimento das contribui��es estabelecidas nesta cl�usula mediante decis�o da referida Assembl�ia Geral, incorrer� na obrigatoriedade do pagamento de multa, cujo valor corresponder� a 5% (cinco por cento) do total da import�ncia a ser recolhida para o SINPRO, acrescida da parcela correspondente � varia��o da TR ou de outro �ndice que vier a substitu�-la, a partir do dia seguinte ao vencimento, cabendo � <b>MANTENEDORA</b> a integral responsabilidade pela multa e demais comina��es, n�o podendo as mesmas, de forma alguma, incidir sobre os sal�rios dos PROFESSORES.
<br><b>Par�grafo segundo</b> � Eventuais discord�ncias dos PROFESSORES, nos termos do Precedente Normativo n� 74 do TST e da ementa do STF, prolatada nos autos do recurso extraordin�rio n� 220-700-1, RS, em 06 de outubro de 1998 e publicada no DJ, edi��o de 13 de novembro de 1998 e do Ac�rd�o de STF, de 07/11/2000, dever�o ser comunicadas oficialmente pelo pr�prio <b>PROFESSOR</b> ao SINPRO, no prazo de 10 dias antes da efetiva��o do primeiro pagamento, j� reajustado, com c�pia � MANTENEDORA, sob pena de perderem efic�cia.
<br><b>Par�grafo terceiro</b> � O SINPRO encaminhar� em tempo h�bil ao SEMESP, ata da assembl�ia geral que fixou a contribui��o, os respectivos valores e a �poca do desconto e do recolhimento.

<tr><td class=titulo>56. N�CLEO INTERSINDICAL DE CONCILIA��O TRABALHISTA
<tr><td class=campo style="text-align:justify">Fica mantido o N�cleo Intersindical de Concilia��o Trabalhista, nos termos previstos pelo artigo 625-C da Consolida��o das Leis do Trabalho, com reda��o dada pela Lei 9.958, de 12 de janeiro de 2000.
<br><b>Par�grafo �nico</b> � O N�cleo Intersindical de Concilia��o Trabalhista ter� suas normas definidas pelo SINPRO-SP e pelo SEMESP e fixadas, sob forma de aditamento, � presente Conven��o Coletiva.

<tr><td class=campo style="text-align:justify">E por estarem justos e acertados, assinam a presente Conven��o Coletiva de Trabalho, a qual ser� depositada na Delegacia Regional do Trabalho de S�o Paulo, nos termos do artigo 614 e par�grafos, para fins de arquivo, de modo a surtir, de imediato, os seus efeitos legais.

<tr><td class=campo style="text-align:justify">S�o Paulo, de junho de 2005. 
<br>
<pre>
<br>Hermes Ferreira Figueiredo              Augusto Cezar Casseb
<br>Presidente do SEMESP                    Presidente do SEMESP S�o Jos� do Rio Preto
<br>
<br>Celso Napolitano                        Luiz Antonio Barbagli
<br>Presidente da FEPESP                    Presidente do SINPRO � S�O PAULO
<br>
<br>Solange Cristina Silva Neisy            Martins de Oliveira Cardoso
<br>Presidente do SINPRO � OSASCO           Presidente do SINPRO � JUNDIA�
<br>
<br>Idelfonso Paz Dias                      Eduardo de Oliveira
<br>Presidente do SINPRO � SANTOS           Presidente do SINPRO � GUARULHOS 
</pre>

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