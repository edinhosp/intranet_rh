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
<title>Conven��o Coletiva 2008/9 - Professores</title>
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
<tr><td class=titulo align="center">CONVEN��O COLETIVA DE TRABALHO PARA 2008/2009
<tr><td class=titulo align="center">SEMESP
<tr><td class=titulo align="center">PROFESSORES 
<tr><td class=campo style="text-align:justify">Entre as partes, de um lado, o Sindicato dos Professores de S�o Paulo; Sindicato dos Professores de Campinas e Regi�o (Americana, Amparo, Araras, Campinas, Esp�rito Santo do Pinhal, Limeira, Mogi-Mirim, Piracicaba e Santa B�rbara D�Oeste); Sindicato dos Professores de Santo Andr�, S�o Bernardo do Campo e S�o Caetano do Sul � SINPRO ABC; Sindicato dos Professores de Santos e Regi�o (Bertioga, Canan�ia, Caraguatatuba, Cubat�o, Eldorado, Guaruj�, Iguape, Ilha Bela, Itanha�m, Itariri, Jacupiranga, Juqui�, Miracatu, Mongagu�, Pariquera-A�u, Pedro de Toledo, Peru�be, Praia Grande, Registro, Santos, S�o Sebasti�o, S�o Vicente, Sete Barras e Ubatuba); Sindicato dos Professores de <b>Osasco</b> e Regi�o (Barueri, Carapicu�ba Cotia e Osasco); Sindicato dos Professores de Jundia�; Sindicato dos Professores de Guarulhos; Sindicato dos Professores de Valinhos e Vinhedo; Sindicato dos Professores de Ja�; Sindicato dos Professores de Indaiatuba, Salto e Itu � SINPRO Vales; Sindicato dos Professores de Sorocaba e Regi�o (Alambari, Alum�nio, Angatuba, Apia�, Ara�ariguama, Arai�oaba da Serra, Bar�o de Antonina, Barra do Chap�u, Bofete, Bom Sucesso de Itarar�, Buri, Campina do Monte Alegre, Cap�o Redondo, Ces�rio Lange, Conchas, Coronel Macedo, Guapiara, Guare�, Ibi�na, Ipero, Itaber�, Ita�, Ita�ca, Itapetininga, Itapeva, Itapirapu� Paulista, Itaporanga, Itarar�, Mairinque, Nova Campina, Paranapanema, Piedade, Pilar do Sul, Porangaba, Quadra, Riacho Grande, Ribeira, Ribeir�o Branco, Ribeir�o Grande, Riversul, Salto de Pirapora, S�o Miguel Arcanjo, S�o Roque, Sarapu�, Sorocaba, Tapira�, Taquarituba, Taquarivai, Tatu�, Torre de Pedra, Vargem Grande Paulista, Votorantim) e Sindicato dos Professores de Educa��o B�sica, Superior, Profissionalizantes, livres de Mogi Gua�u e Itapira � SINPRO Guapira; e a Federa��o dos Professores do Estado de S�o Paulo � FEPESP, entidades com bases territoriais e representatividades fixadas nas respectivas Cartas Sindicais e no que estabelece o inciso I do artigo 8� da Constitui��o Federal e de outro, o Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de S�o Paulo � SEMESP e SEMESP S�o Jos� do Rio Preto, com representatividade fixada em seus registros sindicais, ao final assinados por seus representantes legais, devidamente autorizados pelas competentes Assembl�ias Gerais das respectivas categorias, fica estabelecida, nos termos do artigo 611 e seguintes da Consolida��o das Leis do Trabalho e do artigo 8�, inciso VI da Constitui��o Federal, a presente CONVEN��O COLETIVA DE TRABALHO:

<tr><td class=titulo>1. Abrang�ncia 
<tr><td class=campo style="text-align:justify">Esta Conven��o abrange a categoria econ�mica dos estabelecimentos particulares de ensino superior no Estado de S�o Paulo, aqui designados como MANTENEDORA e a categoria profissional diferenciada dos professores, aqui designada simplesmente como PROFESSOR. 
<br><b>Par�grafo �nico</b> � A categoria dos PROFESSORES abrange todos aqueles que exercem a atividade docente, independentemente da denomina��o sob a qual a fun��o for exercida. Considera-se atividade docente a fun��o de ministrar aulas. 

<tr><td class=titulo>2. Dura��o 
<tr><td class=campo style="text-align:justify">Esta Conven��o Coletiva de Trabalho ter� dura��o de dois anos, com vig�ncia de 1� de mar�o de 2008 a 28 de fevereiro de 2010. 
<br><b>Par�grafo �nico</b> � As cl�usulas poder�o ser reexaminadas na pr�xima data base, em 1� de mar�o de 2009, em virtude de problemas surgidos na sua aplica��o ou do surgimento de normas legais a elas pertinentes, ou em decorr�ncia de aprova��o das propostas apresentadas pela Comiss�o de Aprimoramento das Rela��es de Trabalho prevista na cl�usula 57 da presente Conven��o. 

<tr><td class=titulo>3. Reajuste salarial em 2008 
<tr><td class=campo style="text-align:justify">I. Abril de 2008 � A partir de 1� de abril de 2008, ser� aplicado o reajuste de 4,66% (quatro v�rgula sessenta e seis por cento), sobre os sal�rios devidos em 1� de agosto de 2007. Tal reajuste, referente ao m�s de abril, dever� ser pago at� o 5� dia �til do m�s de junho, juntamente com os sal�rios referentes ao m�s de maio. 
II. Julho de 2008 � Em 1� de julho de 2008, as MANTENEDORAS dever�o aplicar o reajuste de 5,5% (cinco e meio por cento), sobre os sal�rios devidos em 1� de agosto de 2007. 
<br><b>Par�grafo primeiro</b> � Fica estabelecido que o sal�rio de 1� de julho de 2008, reajustado pelo �ndice definido nesta cl�usula, servir� como base de c�lculo para a data base de 1� de mar�o de 2009. 

<tr><td class=titulo>4. Reajuste salarial em 1� de mar�o de 2009 
<tr><td class=campo style="text-align:justify">Em 1� de mar�o de 2009, as MANTENEDORAS dever�o aplicar sobre os sal�rios devidos em 1� de julho de 2008, o percentual definido pela m�dia aritm�tica dos �ndices inflacion�rios do per�odo compreendido entre 1� de mar�o de 2008 e 28 de fevereiro de 2009, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV), composto com 1,20% (um v�rgula vinte por cento). 
<br><b>Par�grafo primeiro</b> � O SEMESP, o SINPRO e a FEPESP comprometem-se a divulgar, em comunicado conjunto, at� 20 de mar�o de 2009, o percentual de reajuste salarial calculado pela f�rmula definida no caput. 
<br><b>Par�grafo segundo</b> � A base de c�lculo para a data-base de 1� de mar�o de 2010 ser� constitu�da pelos sal�rios devidos em 1� de julho de 2008, reajustados em 2009 pela f�rmula definida no caput. 

<tr><td class=titulo>5. Compensa��es salariais 
<tr><td class=campo style="text-align:justify">No ano de 2008 ser� permitida a compensa��o de eventuais antecipa��es salariais concedidas no per�odo compreendido entre 1� de mar�o de 2007 e 28 de fevereiro de 2008. Relativamente � data-base de mar�o de 2009 ser� permitida a compensa��o de eventuais antecipa��es salariais concedidas no per�odo compreendido entre 1� de mar�o de 2008 e 28 de fevereiro de 2009.
<br><b>Par�grafo �nico</b> � N�o ser� permitida, em ambos os casos, a compensa��o daquelas antecipa��es salariais que decorrerem de promo��es, transfer�ncias, ascens�o em plano de carreira e os reajustes concedidos com cl�usula expressa de n�o�compensa��o. 

<tr><td class=titulo>6. Sal�rio do professor ingressante na mantenedora 
<tr><td class=campo style="text-align:justify">A MANTENEDORA n�o poder� contratar nenhum PROFESSOR por sal�rio inferior ao limite salarial m�nimo dos PROFESSORES mais antigos que possuam o mesmo grau de qualifica��o ou titula��o de quem est� sendo contratado, respeitado o quadro de carreira da MANTENEDORA. 
<br><b>Par�grafo �nico</b> � Ao PROFESSOR admitido ap�s 1� de mar�o de 2008 e ap�s 1� de mar�o de 2009, ser�o concedidos os mesmos percentuais de reajustes e aumentos salariais estabelecidos nas cl�usulas 3 e 4, respectivamente, desta norma coletiva. 

<tr><td class=titulo>7. Comprovante de pagamento 
<tr><td class=campo style="text-align:justify">A MANTENEDORA dever� fornecer ao PROFESSOR, mensalmente, comprovante de pagamento, devendo estar discriminados: 
<blockquote style="margin-top:0;margin-bottom:0">a) identifica��o da MANTENEDORA e do estabelecimento de ensino; 
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
<br>m) outros descontos. </blockquote>

<tr><td class=titulo>8. Hora-atividade 
<tr><td class=campo style="text-align:justify">Fica mantido o adicional de 5% (cinco por cento) a t�tulo de hora-atividade, destinado exclusivamente ao pagamento do tempo gasto pelo PROFESSOR, fora do estabelecimento de ensino, na prepara��o de aulas, provas e exerc�cios, bem como na corre��o dos mesmos. 

<tr><td class=titulo>9. Adicional noturno 
<tr><td class=campo style="text-align:justify">O trabalho noturno deve ser pago nas atividades realizadas ap�s as 22 horas e corresponde a 25% (vinte e cinco por cento) do valor da hora-aula. 

<tr><td class=titulo>10. Horas extras 
<tr><td class=campo style="text-align:justify">Considera-se atividade extra todo trabalho desenvolvido em hor�rio diferente daquele habitualmente realizado na semana. As atividades extras devem ser pagas com adicional de 100% (cem por cento). 
<br><b>Par�grafo primeiro</b> � N�o � considerada atividade extra a participa��o em cursos de capacita��o e aperfei�oamento docente, desde que aceita livremente pelo PROFESSOR. 
<br><b>Par�grafo segundo</b> � Ser�o pagas apenas como aulas normais, acrescidas do DSR e da hora-atividade, aquelas que forem adicionadas provisoriamente � carga hor�ria habitual, decorrentes: 
<blockquote style="margin-top:0;margin-bottom:0">a) da substitui��o tempor�ria de um outro PROFESSOR, com dura��o predeterminada, decorrente de licen�a m�dica, maternidade ou para estudos. Nestes casos, a substitui��o dever� ser formalizada atrav�s de documento firmado entre a MANTENEDORA e o PROFESSOR que aceitar realiz�-la; 
<br>b) de substitui��es eventuais de faltas de PROFESSOR respons�vel, desde que aceitas livremente pelo PROFESSOR substituto;
<br>c) de reposi��o de eventuais faltas que foram descontadas dos sal�rios nos meses em que ocorreram; 
<br>d) da realiza��o de cursos eventuais ou de curta dura��o, inclusive cursos de depend�ncia, e aceitas livremente, mediante documento firmado entre o PROFESSOR convidado a ministr�-los e a MANTENEDORA. 
<br>e) do comparecimento a reuni�es did�tico-pedag�gicas, de avalia��o e de planejamento, quando realizadas fora de seu hor�rio habitual de trabalho, desde que aceito livremente pelo PROFESSOR. 
</blockquote>
<b>Par�grafo terceiro</b> � A participa��o em Comiss�es Internas e Externas da Unidade de Ensino da MANTENEDORA, desde que aceita livremente pelo PROFESSOR mediante documento firmado, ser� remunerada como aula ou hora normal, acrescida de DSR. 

<tr><td class=titulo>11. Janelas 
<tr><td class=campo style="text-align:justify">Considera-se janela a aula vaga existente no hor�rio do PROFESSOR entre duas outras aulas ministradas no mesmo turno. O pagamento da janela � obrigat�rio, devendo o PROFESSOR permanecer � disposi��o da MANTENEDORA neste per�odo, ressalvada a aceita��o pelo PROFESSOR, atrav�s de acordo formalizado entre as partes antes do in�cio das aulas, quando as janelas n�o ser�o pagas. 
<br><b>Par�grafo �nico</b> - Ocorrendo a hip�tese da ressalva supra e caso o PROFESSOR seja solicitado esporadicamente a ministrar aulas ou a desenvolver qualquer outra atividade inerente ao magist�rio, no hor�rio de janelas n�o-pagas, essas atividades ser�o remuneradas como aulas extras, com adicional de 100% (cem por cento). 

<tr><td class=titulo>12. Adicional por atividades em outros munic�pios 
<tr><td class=campo style="text-align:justify">Quando o PROFESSOR desenvolver suas atividades a servi�o da mesma MANTENEDORA em munic�pio diferente daquele onde foi contratado e onde ocorre a presta��o habitual do trabalho, dever� receber um adicional de 25% (vinte e cinco por cento) sobre o total de sua remunera��o no novo munic�pio. Quando o PROFESSOR voltar a prestar servi�os no munic�pio de origem, cessar� a obriga��o no pagamento do adicional. 
<br><b>Par�grafo primeiro</b> - Nos casos em que ocorrer a transfer�ncia definitiva do PROFESSOR, aceita livremente por este, em documento firmado entre as partes, n�o haver� a incid�ncia do adicional referido no caput, obrigando-se a MANTENEDORA a efetuar o pagamento de um �nico sal�rio mensal integral, ao PROFESSOR, no ato da transfer�ncia, a t�tulo de ajuda de custo. 
<br><b>Par�grafo segundo</b> - Fica assegurada a garantia de emprego pelo per�odo de seis meses ao PROFESSOR transferido de munic�pio, contados a partir do in�cio do trabalho e/ou da efetiva��o da transfer�ncia. 
<br><b>Par�grafo terceiro</b> � Caso a MANTENEDORA desenvolva atividade acad�mica em munic�pios considerados conurbados, poder� solicitar isen��o do pagamento do adicional determinado no caput, desde que encaminhe material comprobat�rio ao SEMESP, para an�lise e delibera��o do Foro Conciliat�rio para Solu��o de Conflitos Coletivos, previsto na cl�usula 47 da presente Conven��o.

<tr><td class=titulo>13. Composi��o do sal�rio mensal do professor 
<tr><td class=campo style="text-align:justify">O sal�rio do PROFESSOR � composto, no m�nimo, por tr�s itens: o sal�rio base, o descanso semanal remunerado (DSR) e a hora-atividade. O sal�rio base � calculado pela seguinte equa��o: n�mero de aulas semanais multiplicado por 4,5 semanas e multiplicado, ainda, pelo valor da hora-aula (artigo 320, par�grafo 1� da CLT). O DSR corresponde a 1/6 (um sexto) do sal�rio base, acrescido, quando houver, do total de horas extras e do adicional noturno (Lei 605/49). A hora-atividade corresponde a 5% (cinco por cento) do total obtido com a somat�ria de todos os valores acima referidos. 
<br><b>Par�grafo �nico</b> - A remunera��o adicional do PROFESSOR pelo exerc�cio concomitante de fun��o n�o-docente obedecer� aos crit�rios estabelecidos entre a MANTENEDORA e o PROFESSOR que aceitar o cargo. 

<tr><td class=titulo>14. Dura��o da hora-aula 
<tr><td class=campo style="text-align:justify">A dura��o da hora-aula poder� ser de, no m�ximo, cinq�enta minutos. 
<br><b>Par�grafo primeiro</b> � Como exce��o ao disposto no caput, a hora-aula poder� ter a dura��o de sessenta minutos nos cursos tecnol�gicos, desde que tenham sido autorizados ou reconhecidos com essa determina��o expressa e cujos PROFESSORES desses cursos tenham sido contratados nessa condi��o. 
<br><b>Par�grafo segundo</b> � As MANTENEDORAS de Institui��es de Ensino que possuam cursos tecnol�gicos nas condi��es definidas no par�grafo 1� desta cl�usula dever�o apresentar � Comiss�o Permanente de Negocia��o definida na cl�usula 48 da presente Conven��o, at� o dia 15 de agosto de 2008, a documenta��o de autoriza��o ou reconhecimento do curso com a determina��o expressa de hora-aula com dura��o de sessenta minutos sob pena de, em n�o o fazendo, estar sujeita � majora��o do valor do sal�rio-aula de acordo com o que estabelece o par�grafo 4� desta cl�usula. 
<br><b>Par�grafo terceiro</b> � Caso a Comiss�o Permanente de Acompanhamento delibere n�o ter havido determina��o expressa do Minist�rio da Educa��o para que a dura��o da hora-aula dos cursos tecnol�gicos seja de sessenta minutos, a MANTENEDORA dever� majorar o sal�rio-aula de acordo com o que estabelece o par�grafo 4� desta cl�usula. 
<br><b>Par�grafo quarto</b> � Em caso de amplia��o da dura��o da hora-aula vigente, respeitado o limite previsto no caput desta cl�usula, a MANTENEDORA dever� acrescer ao sal�rio-aula j� pago, valor proporcional ao acr�scimo do trabalho. 

<tr><td class=titulo>15. Carga hor�ria 
<tr><td class=campo style="text-align:justify">Quando a MANTENEDORA e o PROFESSOR contratarem carga di�ria de aulas superior aos limites previstos no artigo 318 da CLT, o excedente � carga hor�ria legal ser� remunerado como aula normal, acrescido de DSR, hora-atividade e vantagens pessoais. 
<br><b>Par�grafo �nico</b> � Poder� ser flexibilizada a carga hor�ria do PROFESSOR entre jornadas, no exerc�cio de sua fun��o docente e concomitantemente com a atividade administrativa, n�o havendo assim pagamento, no intervalo, de horas aulas e sal�rios, quando o professor n�o tenha trabalhado no referido intervalo. 

<tr><td class=titulo>16. Prazo para pagamento de sal�rios
<tr><td class=campo style="text-align:justify">Os sal�rios dever�o ser pagos, no m�ximo, at� o quinto dia �til do m�s subseq�ente ao trabalhado. 
<br><b>Par�grafo �nico</b> - O n�o-pagamento dos sal�rios no prazo obriga a MANTENEDORA a pagar multa di�ria, em favor do PROFESSOR, no valor de 1/50 (um cinq�enta avos) de seu sal�rio mensal. 

<tr><td class=titulo>17. Desconto de faltas 
<tr><td class=campo style="text-align:justify">Na ocorr�ncia de faltas, a MANTENEDORA poder� descontar do sal�rio do PROFESSOR, no m�ximo, o n�mero de aulas em que o mesmo esteve ausente, o DSR (1/6), a hora-atividade e demais vantagens pessoais proporcionais a estas aulas. 
<br><b>Par�grafo �nico</b> - � da compet�ncia e de integral responsabilidade da MANTENEDORA estabelecer mecanismos de controle de faltas e de pontualidade dos PROFESSORES, conforme a legisla��o vigente. 

<tr><td class=titulo>18. Atestados m�dicos e abono de faltas 
<tr><td class=campo style="text-align:justify">A MANTENEDORA est� obrigada a abonar as faltas dos PROFESSORES, mediante a apresenta��o de atestados m�dicos ou odontol�gicos. 

<tr><td class=titulo>19. Anota��es na carteira de trabalho 
<tr><td class=campo style="text-align:justify">A MANTENEDORA est� obrigada a promover, em quarenta e oito horas, as anota��es nas Carteiras de Trabalho de seus PROFESSORES, ressalvados eventuais prazos mais amplos permitidos por lei. 
<br><b>Par�grafo �nico</b> - � obrigat�ria a anota��o na Carteira de Trabalho das mudan�as provocadas por ascens�o ou altera��o de titula��o, decorrentes e previstas em plano de carreira. 

<tr><td class=titulo>20. Mudan�a de disciplina 
<tr><td class=campo style="text-align:justify">O PROFESSOR n�o poder� ser transferido de uma disciplina para outra, salvo com seu consentimento expresso e por escrito, sob pena de nulidade da referida transfer�ncia. 

<tr><td class=titulo>21. Redu��o de carga hor�ria por extin��o ou supress�o de disciplina, classe ou turma 
<tr><td class=campo style="text-align:justify">Ocorrendo supress�o de disciplina, classe ou turma, em virtude de altera��o na estrutura curricular prevista ou autorizada pela legisla��o vigente ou por dispositivo regimental devidamente aprovado por �rg�o colegiado da Institui��o de Ensino, o PROFESSOR da disciplina classe ou turma dever� ser comunicado da redu��o da sua carga hor�ria, por escrito, com anteced�ncia m�nima de 30 (trinta) dias do in�cio do per�odo letivo e ter� prioridade para preenchimento de vaga existente em outra classe ou turma ou em outra disciplina para a qual possua habilita��o legal. 
<br><b>Par�grafo primeiro</b> � O PROFESSOR dever� manifestar por escrito, no prazo m�ximo de 5 (cinco) dias ap�s a comunica��o da MANTENEDORA, a n�o-aceita��o da transfer�ncia de disciplina ou de classe ou turma ou da redu��o parcial de sua carga hor�ria. A aus�ncia de manifesta��o do PROFESSOR caracterizar� a sua aceita��o. 
<br><b>Par�grafo segundo</b> � Caso o PROFESSOR n�o aceite a transfer�ncia para outra disciplina, classe ou turma ou a redu��o parcial de carga hor�ria, a MANTENEDORA dever� manter a carga hor�ria semanal existente ou, em caso contr�rio, proceder � rescis�o do contrato de trabalho, por demiss�o sem justa causa.

<tr><td class=titulo>22. Redu��o de carga hor�ria por diminui��o do n�mero de alunos matriculados 
<tr><td class=campo style="text-align:justify">Na ocorr�ncia de diminui��o do n�mero de alunos matriculados que venha a caracterizar a supress�o de turmas, curso ou disciplina, o PROFESSOR do curso em quest�o dever� ser comunicado, por escrito, da redu��o parcial ou total de sua carga hor�ria at� o final da segunda semana de aulas do per�odo letivo. 
<br><b>Par�grafo primeiro</b> - O PROFESSOR dever� manifestar, tamb�m por escrito, a aceita��o ou n�o da redu��o parcial de carga hor�ria no prazo m�ximo de cinco dias ap�s a comunica��o da MANTENEDORA. A aus�ncia de manifesta��o do PROFESSOR caracterizar� a sua n�o-aceita��o. 
<br><b>Par�grafo segundo</b> - Caso o PROFESSOR aceite a redu��o parcial de carga hor�ria, dever� formalizar documento junto � MANTENEDORA e, em n�o aceitando, a MANTENEDORA dever� proceder � rescis�o do contrato de trabalho, por demiss�o sem justa causa, caso seja mantida a redu��o parcial de carga hor�ria. 
<br><b>Par�grafo terceiro</b> - Na hip�tese de rescis�o contratual, por demiss�o sem justa causa, o aviso pr�vio ser� indenizado, estando a MANTENEDORA desobrigada do pagamento do disposto na cl�usula 29 da presente Conven��o � Garantia Semestral de Sal�rios. 
<br><b>Par�grafo quarto</b> - N�o ocorrendo redu��o do n�mero de alunos matriculados que venha a caracterizar supress�o do curso, de turma ou de disciplina, a MANTENEDORA que reduzir a carga hor�ria do PROFESSOR estar� sujeita ao disposto na cl�usula 29 desta Conven��o � Garantia Semestral de Sal�rios � quando ocorrer a rescis�o do contrato de trabalho do PROFESSOR. 

<tr><td class=titulo>23. Abono de faltas por casamento ou luto 
<tr><td class=campo style="text-align:justify">N�o ser�o descontadas, no curso de nove dias corridos, as faltas do PROFESSOR, por motivo de gala ou luto, este em decorr�ncia de falecimento de pai, m�e, filho, c�njuge, companheira (o) e dependente juridicamente reconhecido. 
<br><b>Par�grafo �nico</b> � N�o ser�o descontadas, no curso de tr�s dias, as faltas do PROFESSOR por motivo de falecimento de sogra, sogro, neto, neta, irm�o ou irm�o. 

<tr><td class=titulo>24. Irredutibilidade salarial 
<tr><td class=campo style="text-align:justify">� proibida a redu��o de remunera��o mensal ou de carga hor�ria, ressalvada a ocorr�ncia do disposto nas cl�usulas 21 e 22 da presente Conven��o, ou ainda, quando ocorrer iniciativa expressa do PROFESSOR. Em qualquer hip�tese, � obrigat�ria a concord�ncia rec�proca, firmada por escrito. 
<br><b>Par�grafo primeiro</b> � N�o havendo concord�ncia rec�proca, a parte que deu origem � redu��o prevista nesta cl�usula arcar� com a responsabilidade da rescis�o contratual. 
<br><b>Par�grafo segundo</b> � Outras atividades, ainda que inerentes ao trabalho docente, que n�o sejam as de ministrar aulas, de dura��o tempor�ria e determinada, poder�o ser regulamentadas por contrato entre as partes, contendo a caracteriza��o da atividade, o in�cio e a previs�o do t�rmino. 
<br><b>Par�grafo terceiro</b> � A MANTENEDORA n�o poder� reduzir o valor da hora-aula dos contratos de trabalho vigentes, ainda que venha a instituir ou modificar plano de carreira. 

<tr><td class=titulo>25. Uniformes 
<tr><td class=campo style="text-align:justify">A MANTENEDORA dever� fornecer gratuitamente dois uniformes por ano, quando o seu uso for exigido.

<tr><td class=titulo>26. Licen�a sem remunera��o 
<tr><td class=campo style="text-align:justify">O PROFESSOR com mais de cinco anos ininterruptos de servi�o na MANTENEDORA ter� direito a licenciar-se, sem direito � remunera��o, por um per�odo m�ximo de dois anos, n�o sendo este per�odo de afastamento computado para contagem de tempo de servi�o ou para qualquer outro efeito, inclusive legal. 
<br><b>Par�grafo primeiro</b> - A licen�a ou sua prorroga��o dever� ser comunicada por escrito, � MANTENEDORA, com anteced�ncia m�nima de noventa dias do per�odo letivo, devendo especificar as datas de in�cio e t�rmino do afastamento. A licen�a s� ter� in�cio a partir da data expressa no comunicado, mantendo-se, at� a�, todas as vantagens contratuais. A inten��o de retorno do PROFESSOR � atividade dever� ser comunicada � MANTENEDORA, no m�nimo, sessenta dias antes do t�rmino do afastamento. 
<br><b>Par�grafo segundo</b> - O t�rmino do afastamento dever� coincidir com o in�cio do per�odo letivo. 
<br><b>Par�grafo terceiro</b> - O PROFESSOR que tenha ou exer�a cargo de confian�a dever�, junto com o comunicado de licen�a, solicitar seu desligamento do cargo a partir do in�cio do per�odo de licen�a. 
<br><b>Par�grafo quarto</b> - Considera-se demission�rio o PROFESSOR que, ao t�rmino do afastamento, n�o retornar �s atividades docentes. 
<br><b>Par�grafo quinto</b> - Ocorrendo a dispensa sem justa causa ao t�rmino da licen�a, o PROFESSOR n�o ter� direito � Garantia Semestral de Sal�rios, prevista na cl�usula 29 da presente Conven��o. 

<tr><td class=titulo>27. Licen�a � professora adotante 
<tr><td class=campo style="text-align:justify">Nos termos da Lei 10421, de 15 de abril de 2002, ser� assegurada licen�a maternidade � PROFESSORA que vier a adotar ou obtiver guarda judicial de crian�as, garantido o emprego no per�odo em que a licen�a for concedida. 

<tr><td class=titulo>28. Licen�a paternidade 
<tr><td class=campo style="text-align:justify">A licen�a paternidade ter� dura��o de cinco dias. 

<tr><td class=titulo>29. Garantia semestral de sal�rios 
<tr><td class=campo style="text-align:justify">Ao PROFESSOR demitido sem justa causa, a MANTENEDORA garantir�: 
<blockquote style="margin-top:0;margin-bottom:0">a) no primeiro semestre, a partir de 1� de janeiro, os sal�rios integrais at� o dia 30 de junho; 
<br>b) no segundo semestre, os sal�rios integrais at� o dia 31 de dezembro, ressalvado o par�grafo 4�. 
</blockquote>
<b>Par�grafo primeiro</b> - N�o ter� direito � Garantia Semestral de Sal�rios o PROFESSOR que, na data da comunica��o da dispensa, contar com menos de 18 (dezoito) meses de servi�o prestado � MANTENEDORA, ressalvado o par�grafo 4� desta cl�usula. 
<br><b>Par�grafo segundo</b> � No caso de demiss�es efetuadas no final do primeiro semestre letivo, para n�o ficar obrigada a pagar ao PROFESSOR os sal�rios do segundo semestre, a MANTENEDORA dever� observar as seguintes disposi��es: 
<blockquote style="margin-top:0;margin-bottom:0">a) com aviso pr�vio a ser trabalhado, a demiss�o dever� ser formalizada com anteced�ncia m�nima de trinta dias do in�cio das f�rias; 
<br>b) sendo o aviso pr�vio indenizado, a demiss�o dever� ser formalizada at� um dia antes do in�cio das f�rias, ainda que as f�rias tenham seu in�cio programado para o m�s de julho, obedecendo ao que disp�e a cl�usula 39 da presente Conven��o.
</blockquote>
<b>Par�grafo terceiro</b> - No caso de demiss�es efetuadas no final do ano letivo, para n�o ficar obrigada a pagar ao PROFESSOR os sal�rios do primeiro semestre do ano seguinte, a MANTENEDORA dever� observar as seguintes disposi��es: 
<blockquote style="margin-top:0;margin-bottom:0">a) com aviso pr�vio a ser trabalhado, a demiss�o dever� ser formalizada com anteced�ncia m�nima de trinta dias do in�cio do recesso escolar; 
<br>b) sendo o aviso pr�vio indenizado, a demiss�o dever� ser formalizada at� um dia antes do in�cio do recesso escolar. 
<br>Os dias de aviso pr�vio que forem indenizados n�o contar�o como tempo de servi�o para efeito do pagamento da Garantia Semestral de Sal�rios, conforme o estabelecido nesta cl�usula. 
</blockquote>
<b>Par�grafo quarto</b> - Quando as demiss�es ocorrerem a partir de 16 de outubro, a MANTENEDORA pagar�, independentemente do tempo de servi�o do PROFESSOR, valor correspondente � remunera��o devida at� o dia 18 de janeiro do ano subseq�ente, inclusive, ressalvados os contratos de experi�ncia e por prazo determinado, estes �ltimos v�lidos somente nos casos de substitui��o tempor�ria, conforme o disposto na al�nea a) do par�grafo 2� da cl�usula 10 da presente Conven��o, n�o sendo devido o pagamento acumulativo de aviso-pr�vio. 
<br><b>Par�grafo quinto</b> � Na vig�ncia da presente Conven��o os PROFESSORES ser�o remunerados a partir da data de in�cio de suas atividades na MANTENEDORA, incluindo o per�odo de planejamento escolar. 
<br><b>Par�grafo sexto</b> - Os sal�rios complementares previstos nesta cl�usula ter�o natureza indenizat�ria, n�o integrando, para nenhum efeito legal, o tempo de servi�o do PROFESSOR. 
<br><b>Par�grafo s�timo</b> - O aviso pr�vio de trinta dias previsto no artigo 487 da CLT j� est� integrado �s indeniza��es tratadas nesta cl�usula. 

<tr><td class=titulo>30. Pedido de demiss�o em final de ano letivo 
<tr><td class=campo style="text-align:justify">O PROFESSOR que, no final do ano letivo, comunicar sua demiss�o at� o dia que antecede o in�cio do recesso escolar, ser� dispensado do cumprimento do aviso pr�vio e ter� direito a receber, como indeniza��o, a remunera��o at� o dia 18 de janeiro do ano subseq�ente, independentemente do tempo de servi�o na MANTENEDORA. 

<tr><td class=titulo>31. Garantia de emprego � gestante 
<tr><td class=campo style="text-align:justify">� proibida a dispensa arbitr�ria ou sem justa causa da PROFESSORA gestante, desde o in�cio da gravidez at� sessenta dias ap�s o t�rmino do afastamento legal. O aviso pr�vio come�ar� a contar a partir do t�rmino do per�odo de estabilidade. 

<tr><td class=titulo>32. Creches 
<tr><td class=campo style="text-align:justify">� obrigat�ria a instala��o de local destinado � guarda de crian�as de at� seis meses, quando a MANTENEDORA mantiver contratada, em jornada integral, pelo menos trinta funcion�rias com idade superior a 16 anos. A manuten��o da creche poder� ser substitu�da pelo pagamento do reembolso-creche, nos termos da legisla��o em vigor (artigo 389, par�grafo 1� da CLT e Portarias MTb n� 3296 de 03.09.86 e n�670, de 27/08/97), ou ainda, a celebra��o de conv�nio com uma entidade reconhecidamente id�nea.

<tr><td class=titulo>33. Garantias ao professor em vias de aposentadoria 
<tr><td class=campo style="text-align:justify">Fica assegurado ao PROFESSOR que, comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria integral por tempo de servi�o ou da aposentadoria por idade, a garantia de emprego durante o per�odo que faltar at� a aquisi��o do direito. 
<br><b>Par�grafo primeiro</b> � A garantia de emprego � devida ao PROFESSOR que estiver contratado pela MANTENEDORA h� pelo menos tr�s anos. 
<br><b>Par�grafo segundo</b> � A comprova��o � MANTENEDORA dever� ser feita mediante a apresenta��o de documento que ateste o tempo de servi�o. Este documento dever� ser emitido por pessoa credenciada junto ao �rg�o previdenci�rio. Se o PROFESSOR depender de documenta��o para realiza��o da contagem, ter� um prazo de trinta dias, a contar da data prevista ou marcada para homologa��o da rescis�o contratual. Comprovada a solicita��o de tal documenta��o, os prazos ser�o prorrogados at� que a mesma seja emitida, assegurando-se, nessa situa��o, o pagamento dos sal�rios pelo prazo m�ximo de cento e vinte dias. 
<br><b>Par�grafo terceiro</b> � O contrato de trabalho do PROFESSOR s� poder� ser rescindido por m�tuo acordo homologado pelo SINPRO ou pedido de demiss�o. 
<br><b>Par�grafo quarto</b> � Havendo acordo formal entre as partes, o PROFESSOR poder� exercer outra fun��o, inerente ao magist�rio, durante o per�odo em que estiver garantido pela estabilidade. 
<br><b>Par�grafo quinto</b> � O aviso pr�vio, em caso de demiss�o sem justa causa, integra o per�odo de estabilidade previsto nesta cl�usula. 
<br><b>Par�grafo sexto</b> � Para garantir a estabilidade prevista nesta cl�usula, o PROFESSOR dever� encaminhar � MANTENEDORA, dentro da prorroga��o prevista no par�grafo 2�, documenta��o que demonstre a tramita��o do processo que atesta o tempo de servi�o. 

<tr><td class=titulo>34. Multa por atraso na homologa��o 
<tr><td class=campo style="text-align:justify">A MANTENEDORA deve pagar as verbas devidas na rescis�o contratual no dia seguinte ao t�rmino do aviso pr�vio, quando trabalhado, ou dez dias ap�s o desligamento, quando houver dispensa do cumprimento de aviso pr�vio. O atraso no pagamento das verbas rescis�rias obrigar� a MANTENEDORA ao pagamento de multa, em favor do PROFESSOR, correspondente a um m�s de sua remunera��o, conforme o disposto no par�grafo 8� do artigo 477 da CLT. A partir do vig�simo dia de atraso da homologa��o da rescis�o, a contar da data estabelecida pela legisla��o para o pagamento das verbas rescis�rias, a MATENEDORA estar� obrigada, ainda, a pagar ao PROFESSOR multa di�ria de 0,2% (dois d�cimos percentuais) do sal�rio mensal. A MANTENEDORA estar� desobrigada de pagar a referida multa quando o atraso da homologa��o vier a ocorrer, comprovadamente, por motivos alheios a sua vontade. 
<br><b>Par�grafo �nico</b> � O SINPRO est� obrigado a fornecer comprovante de comparecimento sempre que a MANTENEDORA se apresentar para homologa��o das rescis�es contratuais e comprovar a convoca��o do PROFESSOR. 

<tr><td class=titulo>35. Demiss�o por justa causa 
<tr><td class=campo style="text-align:justify">Quando houver demiss�o por justa causa, nos termos do art. 482 da CLT, a MANTENEDORA est� obrigada a determinar na carta-aviso o motivo que deu origem � dispensa. Caso contr�rio, fica descaracterizada a justa causa. 

<tr><td class=titulo>36. Readmiss�o do professor
<tr><td class=campo style="text-align:justify">O PROFESSOR que for readmitido at� doze meses ap�s o seu desligamento ficar� desobrigado de firmar contrato de experi�ncia. 

<tr><td class=titulo>37. Indeniza��es por dispensa imotivada 
<tr><td class=campo style="text-align:justify">O PROFESSOR demitido sem justa causa ter� direito a uma indeniza��o, al�m do aviso pr�vio legal de trinta dias e das indeniza��es previstas na cl�usula 29 desta Conven��o, quando forem devidas, nas condi��es abaixo especificadas: 
<blockquote style="margin-top:0;margin-bottom:0">a) tr�s (03) dias para cada ano trabalhado na MANTENEDORA; 
<br>b) aviso pr�vio adicional de quinze dias, caso o PROFESSOR tenha, no m�nimo, cinq�enta anos de idade e que, � data do desligamento, conte com pelo menos um ano de servi�o na MANTENEDORA. 
</blockquote>
<b>Par�grafo primeiro</b> � N�o ter� direito � indeniza��o assegurada na al�nea a) do caput o PROFESSOR que tiver recebido, durante pelo menos um ano, pagamento mensal de adicional por tempo de servi�o decorrente de plano de cargos e sal�rios ou de anu�nio, q�inq��nio ou equivalente, cujo valor corresponda a, no m�nimo, 1% (um por cento) do valor da hora-aula por ano trabalhado e, por conseq��ncia, do sal�rio mensal. A MANTENEDORA dever� apresentar, no momento da homologa��o, documentos que comprovem o pagamento ao PROFESSOR do referido adicional por tempo de servi�o. 
<br><b>Par�grafo segundo</b> � N�o ter� direito � indeniza��o assegurada na al�nea b) do caput, o PROFESSOR que, na data de admiss�o na MANTENEDORA, contar com mais de cinq�enta anos de idade. 
<br><b>Par�grafo terceiro</b> � O pagamento das verbas indenizat�rias previstas nesta cl�usula n�o ser� cumulativo, cabendo ao PROFESSOR, no desligamento, o maior valor monet�rio entre os previstos nas al�neas a) e b) do caput. 
<br><b>Par�grafo quarto</b> � Essas indeniza��es n�o contar�o, para nenhum efeito, como tempo de servi�o. 

<tr><td class=titulo>38. Atestados de afastamento e sal�rios 
<tr><td class=campo style="text-align:justify">Sempre que solicitada, a MANTENEDORA dever� fornecer ao PROFESSOR atestado de afastamento e sal�rio (AAS), previsto na legisla��o previdenci�ria. 

<tr><td class=titulo>39. F�rias 
<tr><td class=campo style="text-align:justify">As f�rias anuais dos PROFESSORES ser�o coletivas, com dura��o de trinta dias corridos e gozados em julho de 2008 e julho de 2009. Qualquer altera��o dever� ser aprovada por �rg�o competente, conforme o estabelecido em Estatuto ou Regimento e dever� constar do calend�rio escolar. 
<br><b>Par�grafo primeiro</b> � A MANTENEDORA est� obrigada a pagar o sal�rio das f�rias e o abono constitucional de 1/3 (um ter�o) at� quarenta e oito horas antes do in�cio das f�rias. 
<br><b>Par�grafo segundo</b> � As f�rias n�o poder�o ser iniciadas aos domingos, feriados, dias de compensa��o do descanso semanal remunerado e nem aos s�bados, quando estes n�o forem dias normais de aula. 
<br><b>Par�grafo terceiro</b> � Tamb�m ter� direito �s f�rias coletivas de trinta dias corridos nos per�odos estabelecidos no caput, O PROFESSOR que, al�m de ministrar aulas, tenha cargo de confian�a ou exer�a outras atividades na MANTENEDORA.
Caso o exerc�cio da atividade administrativa impossibilite a concess�o de f�rias nos termos do caput, as f�rias anuais desse PROFESSOR poder�o ser gozadas em dois per�odos, um deles obrigatoriamente no m�s de julho de cada ano. 
<br><b>Par�grafo quarto</b> � Na hip�tese da divis�o das f�rias anuais do PROFESSOR nos termos do par�grafo anterior, um dos per�odos n�o poder� ser inferior a 10 (dez) dias, sendo proibido o exerc�cio de qualquer atividade nesses per�odos. 

<tr><td class=titulo>40. Recesso escolar 
<tr><td class=campo style="text-align:justify">O recesso escolar anual � obrigat�rio e tem dura��o de trinta dias corridos, gozados preferencialmente no m�s de janeiro de cada ano. Durante o recesso escolar anual que n�o pode, de maneira alguma, coincidir com o per�odo definido para as f�rias coletivas do ano respectivo, o PROFESSOR n�o poder� ser convocado para nenhum trabalho. 
<br><b>Par�grafo primeiro</b> � Na vig�ncia da presente Conven��o, as institui��es cujos calend�rios escolares, determinados pelo �rg�o competente conforme o estabelecido em Estatuto ou Regimento, n�o observarem o determinado pelo caput para o recesso escolar anual dos PROFESSORES, poder�o conced�-lo em um per�odo de, no m�nimo vinte dias corridos, e em, no m�ximo, mais dois per�odos com igual n�mero de dias corridos, desde que observem as seguintes condi��es: 
<blockquote style="margin-top:0;margin-bottom:0">a) Vinte dias corridos em janeiro de 2008 e os dois per�odos com igual n�mero de dias corridos, obrigatoriamente no per�odo compreendido entre mar�o de 2008 e fevereiro de 2009. 
<br>b) Vinte dias corridos em janeiro de 2009 e os dois per�odos com igual n�mero de dias corridos, obrigatoriamente no per�odo compreendido entre mar�o de 2009 e fevereiro de 2010. 
</blockquote>
<b>Par�grafo segundo</b> � No caso dos calend�rios escolares preverem a divis�o do recesso escolar dos PROFESSORES, os per�odos definidos na conformidade do par�grafo primeiro n�o poder�o ser iniciados aos domingos, feriados, dias de compensa��o do descanso semanal remunerado e nem aos s�bados, quando estes n�o forem dias normais de aulas. 
<br><b>Par�grafo terceiro</b> � As Institui��es cujas atividades n�o possam ser interrompidas, tais como aquelas desenvolvidas em hospital, cl�nica, laborat�rio de an�lise, escrit�rios experimentais, pesquisas, dentre outros, ou que ministrem cursos em que sejam utilizadas instala��es espec�ficas ou que prestem atendimento � comunidade que n�o pode ser suspenso, poder�o conceder aos PROFESSORES o recesso escolar anual definido no caput de maneira escalonada ao longo de cada ano. 
<br><b>Par�grafo quarto</b> � Os calend�rios escolares que definir�o os per�odos de recesso escolar dos PROFESSORES ser�o obrigatoriamente divulgados aos PROFESSORES at� o in�cio de cada per�odo letivo e enviados ao SINPRO. 

<tr><td class=titulo>41. Delegado representante 
<tr><td class=campo style="text-align:justify">A MANTENEDORA que tiver mais de 50 (cinq�enta) PROFESSORES assegurar� elei��o de Delegados Representantes, com mandato de 1 (um) ano, que ter�o garantia de emprego e sal�rios a partir da inscri��o de sua candidatura at� o t�rmino do semestre letivo em que sua gest�o se encerrar, nos seguintes limites: 
<blockquote style="margin-top:0;margin-bottom:0">a) Na MANTENEDORA que tenha at� 100 (cem) PROFESSORES, ser� garantida a elei��o de 1 (um) delegado representante; 
<br>b) Na MANTENEDORA que tenha mais de 200 (duzentos) PROFESSORES, ser� garantida a elei��o de 2 (dois) delegados representantes; 
</blockquote>
<b>Par�grafo primeiro</b> � O mandato dos Delegados Representantes ser� de um ano. 
<br><b>Par�grafo segundo</b> � A elei��o dos Delegados Representantes ser� realizada pelo SINPRO nas unidades de ensino da MANTENEDORA, por voto direto e secreto. � exigido quorum de 50% (cinq�enta por cento) mais um do corpo docente da unidade onde a elei��o ocorrer. 
<br><b>Par�grafo terceiro</b> � O SINPRO comunicar� a elei��o � MANTENEDORA, com a rela��o dos candidatos inscritos, com anteced�ncia m�nima de sete dias corridos, da data da elei��o. Nenhum candidato poder� ser demitido a partir da data da comunica��o at� o t�rmino da apura��o. 
<br><b>Par�grafo quarto</b> � � condi��o necess�ria que os candidatos sejam filiados ao SINPRO e que tenham, � data da elei��o, pelo menos um ano de servi�o na MANTENEDORA. 

<tr><td class=titulo>42. Quadro de avisos 
<tr><td class=campo style="text-align:justify">A MANTENEDORA dever� colocar, nas salas de professores, quadro de aviso � disposi��o do SINPRO para fixa��o de comunicados de interesse da categoria, sendo vedada a divulga��o de mat�ria pol�tico-partid�ria ou ofensiva a quem quer que seja. 
<br><b>Par�grafo �nico</b> � O dirigente sindical ter� livre acesso � sala dos PROFESSORES, no hor�rio de intervalo das aulas, para atualiza��o do material divulgado no quadro de avisos, uma �nica vez em cada m�s. 

<tr><td class=titulo>43. Assembl�ias sindicais 
<tr><td class=campo style="text-align:justify">Todo PROFESSOR ter� direito a abono de faltas para o comparecimento a assembl�ias da categoria. 
<br><b>Par�grafo primeiro</b> - Na vig�ncia desta Conven��o, os abonos est�o limitados a dois s�bados e mais dois dias �teis para cada per�odo compreendido entre o m�s de mar�o e o m�s de fevereiro do ano subseq�ente. As duas assembl�ias realizadas durante os dias �teis dever�o ocorrer em per�odos distintos. 
<br><b>Par�grafo segundo</b> - O SINPRO ou a FEPESP dever� informar ao SEMESP ou � MANTENEDORA, por escrito, com anteced�ncia m�nima de quinze dias corridos. Na comunica��o dever�o constar a data e o hor�rio da assembl�ia. 
<br><b>Par�grafo terceiro</b> - Os dirigentes sindicais n�o est�o sujeitos ao limite previsto no par�grafo 1� desta cl�usula. As aus�ncias decorrentes do comparecimento �s assembl�ias de suas entidades ser�o abonadas mediante pr�via comunica��o formal � MANTENEDORA. 
<br><b>Par�grafo quarto</b> - A MANTENEDORA poder� exigir dos PROFESSORES e do dirigente sindical atestado emitido pelo SINPRO ou pela FEPESP que comprove o seu comparecimento � assembl�ia. 

<tr><td class=titulo>44. Congressos, simp�sios e equivalentes 
<tr><td class=campo style="text-align:justify">Os abonos de falta para comparecimento a congressos e simp�sios ser�o concedidos mediante aceita��o por parte da MANTENEDORA, que dever� formalizar por escrito a dispensa do PROFESSOR. 
<br><b>Par�grafo �nico</b> - A participa��o do PROFESSOR nos eventos descritos no caput n�o caracterizar� atividade extraordin�ria. 

<tr><td class=titulo>45. Congresso do Sinpro
<tr><td class=campo style="text-align:justify">Em cada ano de vig�ncia desta Conven��o, o SINPRO promover� um evento de natureza pol�tica ou pedag�gica (congresso ou jornada). A MANTENEDORA abonar� as aus�ncias de seus PROFESSORES que participarem do evento, nos seguintes limites: 
<blockquote style="margin-top:0;margin-bottom:0">a) na unidade de ensino que tenha at� 49 (quarenta e nove) PROFESSORES ser� garantido o abono a 1 (um) PROFESSOR; 
<br>b) na unidade de ensino que tenha entre 50 (cinq�enta) e 99 (noventa e nove) PROFESSORES ser� garantido o abono a 2 (dois) PROFESSORES; 
<br>c) na unidade de ensino que tenha mais de 100 (cem) PROFESSORES ser� garantido o abono a 3 (tr�s) PROFESSORES. 
</blockquote>
Tais faltas, limitadas ao m�ximo em dois dias �teis al�m do s�bado, em cada evento, ser�o abonadas mediante a apresenta��o de atestado de comparecimento fornecido pelo SINPRO. O PROFESSOR dever� repor as aulas que, por ventura, sejam necess�rias para complementa��o das horas letivas m�nimas exigidas pela legisla��o. 

<tr><td class=titulo>46. Rela��o nominal 
<tr><td class=campo style="text-align:justify">Na vig�ncia desta Conven��o, obriga-se a MANTENEDORA a encaminhar ao SINPRO, at� o final do m�s de junho de cada ano, a rela��o nominal dos PROFESSORES que integram seu quadro de funcion�rios, acompanhada do valor do sal�rio mensal e das guias das contribui��es sindical e assistencial. 

<tr><td class=titulo>47. Foro Conciliat�rio para Solu��o de Conflitos Coletivos 
<tr><td class=campo style="text-align:justify">Fica mantida a exist�ncia do Foro Conciliat�rio que tem como objetivo procurar resolver quest�es referentes ao n�o cumprimento de normas estabelecidas na presente Conven��o e eventuais diverg�ncias trabalhistas existentes entre a MANTENEDORA e seus PROFESSORES. 
<br><b>Par�grafo primeiro</b> - O Foro ser� composto por membros do SEMESP e do SINPRO. As reuni�es dever�o contar, tamb�m, com as partes em conflito que, se assim o desejarem, poder�o delegar representantes para substitu�-las e/ou serem assistidas por advogados. 
<br><b>Par�grafo segundo</b> - O SEMESP e o SINPRO dever�o indicar os seus representantes no Foro num prazo de trinta dias a contar da assinatura desta Conven��o. 
<br><b>Par�grafo terceiro</b> - Cada se��o do Foro ser� realizada no prazo m�ximo de quinze dias a contar da solicita��o formal e obrigat�ria de qualquer uma das entidades que o comp�em, devendo constar na solicita��o a data, o local e o hor�rio em que a mesma dever� se realizar. O n�o-comparecimento de qualquer uma das partes acarretar� no encerramento imediato das negocia��es. 
<br><b>Par�grafo quarto</b> - Nenhuma das partes envolvidas ingressar� com a��o na Justi�a do Trabalho durante as negocia��es de entendimento. 
<br><b>Par�grafo quinto</b> - Na aus�ncia de solu��o do conflito ou na hip�tese de n�o-comparecimento de qualquer uma das partes, a comiss�o respons�vel pelo Foro fornecer� certid�o atestando o encerramento da negocia��o. 
<br><b>Par�grafo sexto</b> - Na hip�tese de sucesso das negocia��es, a crit�rio do Foro, a MANTENEDORA ficar� desobrigada de arcar com a multa prevista na cl�usula 55 desta Conven��o. 
<br><b>Par�grafo s�timo</b> - As decis�es do Foro ter�o efic�cia legal entre as partes acordantes. O descumprimento das decis�es assumidas gerar� multa a ser estabelecida no Foro, independentemente daquelas j� estabelecidas nesta Conven��o.
<br><b>Par�grafo oitavo</b> � Na hip�tese de incapacidade econ�mico-financeira das MANTENEDORAS, os casos ser�o remetidos para an�lise e delibera��o deste foro. 

<tr><td class=titulo>48. Comiss�o Permanente de Negocia��o 
<tr><td class=campo style="text-align:justify">Fica mantida a Comiss�o Permanente de Negocia��o constitu�da de forma parit�ria, por tr�s representantes das entidades sindicais (profissional e econ�mica), com o objetivo de: 
<blockquote style="margin-top:0;margin-bottom:0">a) fiscalizar o cumprimento das cl�usulas vigentes; 
<br>b) elucidar eventuais diverg�ncias de interpreta��o das cl�usulas desta Conven��o; 
<br>c) discutir quest�es n�o-contempladas na presente Conven��o. 
<br>d) deliberar no prazo m�ximo de trinta dias a contar da data da solicita��o protocolizada no SEMESP, sobre modifica��o de pagamento da assist�ncia m�dico-hospitalar, conforme os par�grafos 1� e 3� da cl�usula 50 da presente Conven��o e sobre o valor da remunera��o da hora-aula, conforme o par�grafo 2� da cl�usula 14 da presente Conven��o. 
<br>e) criar subs�dios para a Comiss�o de Tratativas Salariais, atrav�s da elabora��o de documentos, para a defini��o das fun��es/atividades e o regime de trabalho dos PROFESSORES. 
</blockquote>
<b>Par�grafo primeiro</b> - As entidades sindicais componentes da Comiss�o Permanente de Negocia��o indicar�o seus representantes, no prazo m�ximo de trinta dias corridos, a contar da assinatura da presente Conven��o. 
<br><b>Par�grafo segundo</b> - A Comiss�o Permanente de Negocia��o dever� reunir-se mensalmente, no d�cimo dia �til, �s 15 horas, alternadamente nas sedes das entidades sindicais que a comp�em. No caso espec�fico do item �d� do caput, dever� haver convoca��o espec�fica feita pela entidade sindical patronal. 

<tr><td class=titulo>49. Acordos internos - cl�usulas mais favor�veis 
<tr><td class=campo style="text-align:justify">Ficam assegurados os direitos mais favor�veis decorrentes de acordos internos ou de acordos coletivos de trabalho celebrados entre a MANTENEDORA e o SINPRO. 

<tr><td class=titulo>50. Assist�ncia m�dico-hospitalar 
<tr><td class=campo style="text-align:justify">A MANTENEDORA est� obrigada a assegurar, �s suas expensas, nos limites estabelecidos nesta cl�usula, assist�ncia m�dico-hospitalar a todos os seus PROFESSORES, sendo-lhe facultada a escolha por plano de sa�de, seguro-sa�de ou conv�nios com empresas prestadoras de servi�os m�dico-hospitalares. Poder� ainda prestar a referida assist�ncia diretamente, em se tratando de institui��es que disponham de servi�os de sa�de e hospitais pr�prios ou conveniados. Qualquer que seja a op��o feita, a assist�ncia m�dico-hospitalar deve assegurar as condi��es e os requisitos m�nimos que seguem relacionados: 
<br><b>1.Abrang�ncia </b>
<blockquote style="margin-top:0;margin-bottom:0">A assist�ncia m�dico-hospitalar deve ser realizada no munic�pio onde funciona o estabelecimento de ensino superior ou onde vive o PROFESSOR, a crit�rio da MANTENEDORA. Em casos de emerg�ncia, dever� haver garantia de atendimento integral em qualquer localidade do Estado de S�o Paulo ou fixa��o, em contrato, de formas de reembolso. 
</blockquote>
<b>2. Coberturas m�nimas </b>
<blockquote style="margin-top:0;margin-bottom:0">2.1 Quarto para quatro pacientes, no m�ximo. 
<br>2.2 Consultas.
<br>2.3 Prazo de interna��o de 365 dias por ano (comum e UTI/CTI) 
<br>2.4 Parto, independentemente do estado grav�dico. 
<br>2.5 Mol�stias infecto-contagiosas que exijam interna��o. 
<br>2.6 Exames laboratoriais, ambulatoriais e hospitalares. 
</blockquote>
<b>3. Car�ncia </b>
<blockquote style="margin-top:0;margin-bottom:0">N�o haver� car�ncia na presta��o dos servi�os m�dicos e laboratoriais. 
</blockquote>
<b>4. Professor ingressante </b>
<blockquote style="margin-top:0;margin-bottom:0">N�o haver� car�ncia para o PROFESSOR ingressante, independentemente do m�s em que for contratado. 
</blockquote>
<b>5. Pagamento </b>
<blockquote style="margin-top:0;margin-bottom:0">Caber� ao PROFESSOR o pagamento de 10% (dez por cento) do valor da Assist�ncia M�dica, respeitado o disposto nos par�grafos 1�, 2� e 3�. 
</blockquote>
<b>Par�grafo primeiro</b> � A MANTENEDORA dever� enviar ao SINPRO c�pia do contrato formalizado com a empresa de assist�ncia m�dico�hospitalar ou de seguro sa�de ou de medicina de grupo que comprove o valor pago. 
<br><b>Par�grafo segundo</b> � Caso a assist�ncia m�dico-hospitalar vigente na Institui��o venha a sofrer reajuste em virtude de poss�veis modifica��es estabelecidas em legisla��o que abranja o segmento - Lei 9.656, de 03 de junho de 1998 e MP 2.097-39, de 26 de abril de 2001, ou que vierem a ser estabelecidas em lei, ou por mudan�a de empresa prestadora de servi�o, a pedido dos empregados da Institui��o ou por quebra de contrato, unilateralmente, por parte da atual empresa prestadora de servi�o, a MANTENEDORA continuar� a contribuir com o valor mensal vigente at� a data da modifica��o, devendo o PROFESSOR arcar com o valor excedente, que ser� descontado em folha e consignado no comprovante de pagamento, nos termos do artigo 462 da CLT. 
<br><b>Par�grafo terceiro</b> � Caso ocorra mudan�a de empresa prestadora de servi�o, por decis�o unilateral da MANTENEDORA, com conseq�ente reajuste no valor vigente, o PROFESSOR estar� isento do pagamento do valor excedente, cabendo � MANTENEDORA prover integralmente a assist�ncia m�dico-hospitalar, sem nenhum �nus para o PROFESSOR. 
<br><b>Par�grafo quarto</b> � Para efeito do disposto no par�grafo primeiro desta cl�usula, caber� � MANTENEDORA remeter a documenta��o comprobat�ria para an�lise e delibera��o da Comiss�o Permanente de Negocia��o. 
<br><b>Par�grafo quinto</b> � Fica facultado ao PROFESSOR optar pela presta��o de assist�ncia m�dico-hospitalar em uma �nica institui��o de ensino, quando mantiver mais de um v�nculo empregat�cio como PROFESSOR. � necess�rio que o PROFESSOR se manifeste por escrito, com anteced�ncia m�nima de vinte dias, para que a MANTENEDORA possa proceder � suspens�o dos servi�os.
<br><b>Par�grafo sexto</b> � Caso o PROFESSOR mantenha v�nculo empregat�cio com mais de uma Institui��o de Ensino, as MANTENEDORAS, em conjunto, poder�o optar por conceder-lhe um �nico plano de sa�de, pago por elas, em regime de cotiza��o de custos, respeitadas as condi��es estabelecidas nesta cl�usula. 
<br><b>Par�grafo s�timo</b> � Mediante pagamento complementar e ades�o facultativa, devidamente documentada, o PROFESSOR poder� optar pela amplia��o dos servi�os de sa�de garantidos nesta Conven��o ou estend�-los a seus dependentes. 

<tr><td class=titulo>51. Bolsas de estudo 
<tr><td class=campo style="text-align:justify">Todo PROFESSOR tem direito a bolsas de estudo integrais, incluindo matr�cula, para si, seus filhos ou dependentes legais, estes �ltimos entendidos como aqueles reconhecidos pela legisla��o do Imposto de Renda ou aqueles que estejam sob a guarda judicial do PROFESSOR e vivam sob sua depend�ncia econ�mica, devidamente comprovada. Os filhos do PROFESSOR poder�o usufruir as bolsas de estudo integrais, sem qualquer �nus, desde que n�o tenham vinte e cinco anos completos ou mais na data da efetiva��o da matr�cula no curso superior. As bolsas de estudo s�o v�lidas para cursos de gradua��o, p�s-gradua��o ou seq�enciais existentes e administrados pela MANTENEDORA para a qual o PROFESSOR trabalha, observado o disposto nesta cl�usula e par�grafos seguintes. 
<br><b>Par�grafo primeiro</b> � O direito �s bolsas de estudo passa a vigorar ao t�rmino do contrato de experi�ncia, cuja dura��o n�o pode exceder de 90 (noventa) dias, conforme par�grafo �nico do artigo 445 da CLT. 
<br><b>Par�grafo segundo</b> - A MANTENEDORA est� obrigada a conceder, no m�ximo, duas bolsas de estudo, sendo que, nos cursos de gradua��o ou seq�enciais, n�o ser� poss�vel que o bolsista conclua mais de um curso nessa condi��o. 
<br><b>Par�grafo terceiro</b> � A utiliza��o do benef�cio previsto nesta cl�usula, caracterizada como doa��o por n�o impor qualquer contrapresta��o de servi�os, � transit�ria e n�o habitual e, por isso, n�o possui car�ter remunerat�rio e nem se vincula, para nenhum efeito, ao sal�rio ou remunera��o percebida pelo PROFESSOR, nos termos do inciso XIX, do par�grafo 9� do artigo 214 do Decreto 3.048, de 06 de maio de 1999 e da Lei 10.243, de 19 de junho de 2001 e visa a capacita��o dos benefici�rios. 
<br><b>Par�grafo quarto</b> - As bolsas de estudo ser�o mantidas quando o PROFESSOR estiver licenciado para tratamento de sa�de ou em gozo de licen�a mediante anu�ncia da MANTENEDORA, excetuado o disposto na cl�usula 26 da presente Conven��o � Licen�a sem Remunera��o. 
<br><b>Par�grafo quinto</b> - No caso de falecimento do PROFESSOR, os dependentes que j� se encontram estudando em estabelecimento de ensino superior da MANTENEDORA continuar�o a gozar das bolsas de estudo at� o final do curso, ressalvado o disposto no par�grafo 8� desta cl�usula. 
<br><b>Par�grafo sexto</b> - No caso de dispensa sem justa causa durante o per�odo letivo, ficam garantidas ao PROFESSOR, at� o final do per�odo letivo, as bolsas de estudo j� existentes. 
<br><b>Par�grafo s�timo</b> - As bolsas de estudo integrais em cursos de p�s-gradua��o ou especializa��o existentes e administrados pela MANTENEDORA s�o v�lidas exclusivamente para o PROFESSOR, em �reas correlatas �s disciplinas que o mesmo ministra na Institui��o e que visem a capacita��o docente, respeitados os crit�rios de sele��o exigidos para ingresso no mesmo e obedecer�o as seguintes condi��es :.
<blockquote style="margin-top:0;margin-bottom:0">a) nos cursos stricto sensu ou de especializa��o que fixem um n�mero m�ximo de alunos por turma, s�o limitadas em 30% (trinta por cento) do total de vagas oferecidas; 
<br>b) nos cursos de p�s-gradua��o lato sensu n�o haver� limites de vagas. Caso a estrutura do curso torne necess�ria a limita��o do n�mero de alunos ser� observado o disposto na al�nea �a� deste par�grafo. 
</blockquote>
<b>Par�grafo oitavo</b> � Os bolsistas que forem reprovados no per�odo letivo perder�o o direito � bolsa de estudo, voltando a gozar do benef�cio quando lograrem aprova��o no referido per�odo. As disciplinas cursadas em regime de depend�ncia ser�o de total responsabilidade do bolsista, arcando o mesmo com o seu custo. 
<br><b>Par�grafo nono</b> - Considera-se adquirido o direito daquele PROFESSOR que j� esteja usufruindo bolsas de estudo em n�mero superior ao definido nesta cl�usula. 

<tr><td class=titulo>52. Autoriza��o para desconto em folha de pagamento 
<tr><td class=campo style="text-align:justify">O desconto do professor em folha de pagamento somente poder� ser realizado mediante sua autoriza��o, nos termos dos artigos 462 e 545 da CLT, quando os valores forem destinados ao custeio de pr�mios de seguro, planos de sa�de, mensalidades associativas ou outras que constem da sua expressa autoriza��o, desde que n�o haja previs�o expressa de desconto na presente norma coletiva. 
<br><b>Par�grafo �nico</b> � Encontra-se no SINPRO, � disposi��o da MANTENEDORA, c�pia de autoriza��o do PROFESSOR para o desconto da mensalidade associativa. 

<tr><td class=titulo>53. Estabilidade para portadores de doen�as graves 
<tr><td class=campo style="text-align:justify">Fica assegurada, at� alta m�dica, considerada como apto ao trabalho, ou eventual concess�o de aposentadoria por invalidez, estabilidade no emprego aos PROFESSORES acometidos por doen�as graves ou incur�veis e aos PROFESSORES portadores do v�rus HIV que vierem a apresentar qualquer tipo de infec��o ou doen�a oportunista, resultante da patologia de base. 
<br><b>Par�grafo �nico</b> � S�o consideradas doen�as graves ou incur�veis, a tuberculose ativa, aliena��o mental, esclerose m�ltipla, neoplasia maligna, cegueira definitiva, hansen�ase, cardiopatia grave, doen�a de Parkinson, paralisia irrevers�vel e incapacitante, espondiloastrose anquilosante, neofropatia grave, estados do Mal de Paget (oste�te deformante) e contamina��o grave por radia��o. 

<tr><td class=titulo>54. Garantias ao professor com seq�elas ocasionadas por doen�as profissionais ou acidente de trabalho 
<tr><td class=campo style="text-align:justify">Ser� garantida ao PROFESSOR acidentado no trabalho ou acometido por doen�a profissional a perman�ncia na empresa em fun��o compat�vel com o seu estado f�sico, sem preju�zo na remunera��o antes percebida, desde que, ap�s o acidente ou comprova��o da aquisi��o de doen�a profissional, apresente, cumulativamente, redu��o da capacidade laboral, atestada pelo �rg�o oficial e que se tenha tornado incapaz de exercer a fun��o que anteriormente desempenhava, obrigado, por�m, o PROFESSOR nessa situa��o a participar dos processos de readapta��o e reabilita��o profissional. 
<br><b>Par�grafo �nico</b> � O per�odo de estabilidade do PROFESSOR que se encontre participando dos processos de readapta��o e reabilita��o profissional ser� o previsto em lei. 

<tr><td class=titulo>55. Multa por descumprimento da Conven��o
<tr><td class=campo style="text-align:justify">O descumprimento desta Conven��o obrigar� a MANTENEDORA ao pagamento de multa correspondente a 1% (um por cento) do sal�rio do PROFESSOR, para cada uma das cl�usulas n�o-cumpridas, acrescidas de juros, a cada PROFESSOR prejudicado. 
<br><b>Par�grafo �nico</b> � A MANTENEDORA est� desobrigada de arcar com a multa prevista no caput, caso a cl�usula descumprida j� estabele�a uma multa pelo seu n�o�cumprimento. 

<tr><td class=titulo>56. Disposi��es transit�rias 
<tr><td class=campo style="text-align:justify">Fica mantida a �Comiss�o de Aprimoramento das Rela��es de Trabalho�, composta de forma parit�ria, por quatro membros de cada uma das categorias profissional e econ�mica, indicados pela FEPESP e pelo SEMESP e/ou SEMESP/SJ RIO PRETO, com o objetivo de apresentar, at� 30 de novembro de 2008, proposta de regulamenta��o dos seguintes temas 
<blockquote style="margin-top:0;margin-bottom:0">a) rela��es de trabalho envolvendo aplica��es de novas tecnologias, cursos semipresenciais e cursos modulares e seq�enciais; 
<br>b) planos de carreira das Institui��es privadas de ensino; 
</blockquote>
<b>Par�grafo primeiro</b> � A primeira reuni�o da �Comiss�o de Aprimoramento das Rela��es de Trabalho� ser� realizada �s 10 horas do dia 27 de maio de 2008, na sede da FEPESP, em S�o Paulo, quando ocorrer� a aprova��o do regimento de funcionamento. 
<br><b>Par�grafo segundo</b> � Os estudos, relat�rios e delibera��es da �Comiss�o de Aprimoramento das Rela��es do Trabalho�, ser�o submetidos �s delibera��es das Assembl�ias convocadas pelas respectivas entidades sindicais e, uma vez aprovadas, inclu�das na presente Conven��o, a partir da pr�xima data base, em 1� de mar�o de 2009. 

<tr><td class=campo style="text-align:justify">E por estarem justos e acertados, assinam a presente Conven��o Coletiva de Trabalho, a qual ser� depositada na Delegacia Regional do Trabalho de S�o Paulo, nos termos do artigo 614 e par�grafos, para fins de arquivo, de modo a surtir, de imediato, os seus efeitos legais. 

<tr><td class=campo style="text-align:justify">S�o Paulo, 30 de maio de 2008 

<br>
<pre>
<br>Hermes Ferreira Figueiredo                  Celso Napolitano 
<br>Presidente do SEMESP                        Presidente da FEPESP 
<br>
<br>H�lder Abud Paranhos                        Luiz Antonio Barbagli 
<br>Presidente do SINPRO Sorocaba               Presidente do SINPRO � S�O PAULO 
<br>
<br>Rubens Gon�alves Aniz                       Marco Aur�lio Arruda Aranha 
<br>Presidente do SINPRO � Osasco               Presidente do SINPRO � Salto, Indaiatuba e Itu 
<br>
<br>Martinho Condini                            Reginaldo Alberto Meloni 
<br>Presidente do SINPRO � Jundia�              Presidente do SINPRO � Campinas 
<br>
<br>Alo�sio Alves da Silva                      Rubens Gabriel Abdal 
<br>Presidente do SINPRO � ABC                  Presidente do SINPRO � Valinhos e Vinhedo 
<br>
<br>Ildefonso Paz Dias                          Andr�a Luciana Harada Sousa 
<br>Presidente do SINPRO � Santos               Presidente do SINPRO � Guarulhos 
<br>
<br>M�rcio Silva Sampaio Lopes                  Samuel Cristiano F�vero 
<br>Presidente do SINPRO Mogi Gua�u e Itapira   Presidente do SINPRO - Jau
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

