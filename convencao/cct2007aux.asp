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
<title>Conven��o Coletiva 2007 - Auxiliares</title>
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
<!-- <b>AUXILIARES</b> -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td class=titulo align="center">CONVEN��O COLETIVA DE TRABALHO - 2007</td></tr>
<tr><td class=titulo align="center">AUXILIARES DE ADMINISTRA��O ESCOLAR
<tr><td class=titulo align="center">FETEE ENSINO SUPERIOR
<tr><td class=campo style="text-align:justify">Entre as partes, de um lado, os estabelecimentos de ensino superior, representados pelo Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior do Estado de S�o Paulo � SEMESP, SEMESP S�O JOS� DO RIO PRETO, entidades sindicais de 1� grau, coordenadora e representativa, nos termos do artigo 611, par�grafo 1�, da Consolida��o das Leis do Trabalho, da categoria econ�mica �Entidades Mantenedoras de Estabelecimentos de Ensino Superior�, do 1� grupo � Estabelecimentos de Ensino - do plano da Confedera��o Nacional de Educa��o e Cultura, conforme estabelecido em sua Carta Sindical e de outro, Federa��o dos Trabalhadores em Estabelecimentos de Ensino do Estado de S�o Paulo � FETEE-SP, entidade sindical de 2� grau, coordenadora e representativa da categoria profissional �Auxiliares de Administra��o Escolar (empregados em estabelecimentos de ensino)�, do 1� grupo - Trabalhadores em Estabelecimentos de Ensino - do plano da Confedera��o Nacional dos Trabalhadores em Estabelecimentos de Educa��o e Cultura, nos termos do par�grafo 2�, artigo 611, da Consolida��o das Leis do Trabalho, conforme estabelecido em sua Carta Sindical e Sindicato dos Auxiliares de Administra��o Escolar do ABC (Santo Andr�, S�o Bernardo do Campo e S�o Caetano do Sul); Sindicato dos Professores e Auxiliares Administrativos de Ara�atuba e Regi�o (Ara�atuba e Birigui); Sindicato dos Auxiliares de Administra��o Escolar de Bauru; Sindicato dos Professores e Auxiliares de Administra��o Escolar de Bragan�a Paulista; Sindicato dos Professores e Trabalhadores em Educa��o de Dracena e Regi�o (Junqueir�polis, Monte Castelo, Nova Guataporanga, Ouro Verde, Panorama, Paulic�ia, Santa Mercedes, S�o Jo�o do Pau D�Alho, Tupi Paulista); Sindicato dos Professores e Auxiliares Administrativos de Fernand�polis (Auriflama, Estrela D�Oeste, General Salgado, Ilha Solteira, Nhandeara, Pereira Barreto, Santa F� do Sul, Ur�nia); Sindicato dos Professores e Auxiliares Administrativos de Jales; Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Lins; Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Mar�lia; Sindicato dos Auxiliares de Administra��o Escolar de Mogi das Cruzes; Sindicato dos Auxiliares de Administra��o Escolar de Piracicaba; Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Presidente Prudente; Sindicato dos Professores e Auxiliares de Administra��o Escolar de Ribeir�o Preto; Sindicato dos Auxiliares de Administra��o Escolar de Santos; Sindicato dos Professores e Auxiliares Administrativos de S�o Jos� dos Campos (Ca�apava, Guararema, Igarat�, Jacare�, Jambeiro, Monteiro Lobato, Natividade da Serra, Paraibuna, Reden��o da Serra, Santa Isabel, Santo Antonio do Pinhal, S�o Bento do Sapuca�, S�o Francisco Xavier); Sindicato dos Auxiliares de Administra��o Escolar de Sorocaba; Sindicato dos Professores e Auxiliares Administrativos de Votuporanga; Sindicato dos Auxiliares de Administra��o Escolar de S�o Jos� do Rio Preto, entidades de 1� grau, coordenadoras e representativas da categoria profissional diferenciada �Professores�, do 1� grupo - Trabalhadores em Estabelecimentos de Ensino - do plano da Confedera��o Nacional dos Trabalhadores em Estabelecimentos de Educa��o e Cultura, com sua representatividade fixada na carta sindical ou nos termos dos incisos I e II, artigo 8�, da Constitui��o Federal, por seus representantes legais, ao final assinados, todos devidamente autorizados por suas assembl�ias gerais, fica estabelecida, nos termos do artigo 611 e seguintes da Consolida��o das Leis do Trabalho, do artigo 8�, inciso VI, do artigo 7�, inciso XXVI e artigo 5�, caput e inciso I, todos da Constitui��o Federal, a presente CONVEN��O COLETIVA DE TRABALHO:

<tr><td class=titulo>1. ABRANG�NCIA 
<tr><td class=campo style="text-align:justify">Esta Conven��o Coletiva de Trabalho abrange a categoria profissional �AUXILIARES DE ADMINISTRA��O ESCOLAR� (empregados em estabelecimentos de ensino), do 1� grupo � Trabalhadores em Estabelecimentos de Ensino � do plano da Confedera��o Nacional dos Trabalhadores em Estabelecimentos de Educa��o e Cultura, em dia com as suas obriga��es estatut�rias e das delibera��es da Assembl�ia, doravante designados como �AUXILIARES� e a categoria econ�mica �estabelecimentos de ensino superior do Estado de S�o Paulo�, integrante do 1� grupo � Estabelecimentos de Ensino � do plano da Confedera��o Nacional de Educa��o e Cultura, representados pelo Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de S�o Paulo, doravante designados como �MANTENEDORAS�. 
<br><b>Par�grafo �nico</b> � A categoria profissional dos AUXILIARES DE ADMINISTRA��O ESCOLAR abrange todos aqueles que, sob qualquer t�tulo ou denomina��o, exercem atividades n�o docentes nos estabelecimentos particulares de ensino superior, consoante a representa��o contida em sua Carta Sindical. 

<tr><td class=titulo>2. DURA��O 
<tr><td class=campo style="text-align:justify">Esta Conven��o Coletiva de Trabalho ter� a dura��o de um ano, com vig�ncia de 1� de mar�o de 2007 a 29 de fevereiro de 2008. 
<br><b>Par�grafo �nico</b> � As cl�usulas constantes da presente norma poder�o ser reexaminadas em virtude de problemas surgidos na sua aplica��o ou do surgimento de normas legais a elas pertinentes, para as devidas adequa��es. 

<tr><td class=titulo>3. REAJUSTE SALARIAL 
<tr><td class=campo style="text-align:justify"><b>a) ABRIL 2007</b>
<br>Em 1� de abril de 2007, as MANTENEDORAS dever�o aplicar sobre os sal�rios devidos em 1� de mar�o de 2007, um reajuste salarial de 3,5% (tr�s v�rgula cinco por cento). 
<br><b>Par�grafo primeiro</b> � Os sal�rios reajustados conforme estabelecido no caput desta cl�usula, dever�o ser pagos at� o quinto dia �til do m�s de maio de 2007. 
<tr><td class=campo style="text-align:justify"><b>b) AGOSTO 2007</b>
<br>Em 1� de agosto de 2007, as MANTENEDORAS dever�o aplicar tamb�m sobre os sal�rios devidos em 1� de mar�o de 2007, um reajuste salarial de 4% (quatro por cento). 
<br><b>Par�grafo primeiro</b> � Os sal�rios reajustados conforme estabelecido no caput desta cl�usula, dever�o ser pagos at� o quinto dia �til do m�s de setembro de 2007. 
<br><b>Par�grafo segundo</b> � O sal�rio de agosto de 2007 ser� a base de c�lculo para a data base da Conven��o Coletiva de Trabalho de 2008.

<tr><td class=titulo>4. COMPENSA��ES SALARIAIS 
<tr><td class=campo style="text-align:justify">Ser� permitida a compensa��o de eventuais antecipa��es de reajustes salariais concedidas no per�odo de vig�ncia da Conven��o 2006/07 relativa ao per�odo de 1� de mar�o de 2006 a 28 de fevereiro de 2007, exceto o previsto na cl�usula 3� da presente conven��o e os reajustes que decorrerem de promo��es, transfer�ncias, ascens�o em plano de carreira al�m daqueles reajustes espont�neos. 

<tr><td class=titulo>5. SAL�RIO DO AUXILIAR INGRESSANTE NA MANTENEDORA 
<tr><td class=campo style="text-align:justify">A MANTENEDORA n�o poder� contratar nenhum AUXILIAR por sal�rio inferior ao limite salarial m�nimo dos AUXILIARES mais antigos que possuam o mesmo grau de qualifica��o ou titula��o de quem est� sendo contratado, respeitado o quadro de carreira da MANTENEDORA. 
<br><b>Par�grafo �nico</b> � Ao AUXILIAR admitido ap�s 1� de mar�o de 2007, ser�o concedidos os mesmos percentuais de reajustes e aumentos salariais estabelecidos nesta norma coletiva. 

<tr><td class=titulo>6. PRAZO E FORMA DE PAGAMENTO DOS SAL�RIOS 
<tr><td class=campo style="text-align:justify">Os sal�rios dever�o ser pagos, no m�ximo, at� o 5� dia �til do m�s subseq�ente ao trabalhado. 
<br><b>Par�grafo primeiro</b> � O n�o pagamento dos sal�rios no prazo obriga a MANTENEDORA a pagar multa di�ria, em favor do AUXILIAR, no valor de 1/30 (um trinta avos) de seu sal�rio mensal. 
<br><b>Par�grafo segundo</b> � As MANTENEDORAS que n�o efetuarem o pagamento dos sal�rios em moeda corrente dever�o proporcionar aos AUXILIARES tempo h�bil para o recebimento no banco ou no posto banc�rio, excluindo-se o hor�rio de refei��o. 
<br><b>Par�grafo terceiro</b> � As MANTENEDORAS que eventualmente alegarem impossibilidade de cumprimento do prazo estabelecido no par�grafo anterior, poder�o requerer ao Foro Conciliat�rio outra data de pagamento de sal�rios, desde que n�o ultrapasse o d�cimo dia do m�s, ficando sujeitas �s decis�es adotadas no mesmo. 

<tr><td class=titulo>7. COMPROVANTES DE PAGAMENTO 
<tr><td class=campo style="text-align:justify">A MANTENEDORA dever� fornecer ao AUXILIAR, mensalmente, comprovante de pagamento, devendo estar discriminados, quando for o caso: 
<blockquote style="margin-top:0;margin-bottom:0">a) identifica��o da MANTENEDORA e do Estabelecimento de Ensino; 
<br>b) identifica��o do AUXILIAR; 
<br>c) denomina��o da fun��o, se houver faixas salariais diferenciadas;
<br>d) carga hor�ria mensal; 
<br>e) outros eventuais adicionais; 
<br>f) descanso semanal remunerado; 
<br>g) horas extras realizadas; 
<br>h) valor do recolhimento do FGTS; 
<br>i) desconto previdenci�rio; 
<br>j) outros descontos.
</blockquote>
<tr><td class=titulo>8. ADICIONAL NOTURNO 
<tr><td class=campo style="text-align:justify">O adicional noturno deve ser pago nas atividades realizadas ap�s as 22 horas e corresponde a 25% (vinte e cinco por cento) do valor das horas trabalhadas. 

<tr><td class=titulo>9. HORAS EXTRAS 
<tr><td class=campo style="text-align:justify">Considera-se atividade extra todo trabalho desenvolvido em hor�rio diferente daquele habitualmente realizado na semana. As tr�s primeiras horas extras semanais devem ser pagas com o adicional de 50% (cinq�enta por cento) e as seguintes, com o adicional de 100% (cem por cento). 
<br><b>Par�grafo primeiro</b> � Caso a MANTENEDORA implante o sistema de Banco de Horas dever� ser observado o disposto na cl�usula pr�pria que regula a mat�ria, integrante da presente norma coletiva. 
<br><b>Par�grafo segundo</b> � Exceto nas hip�teses de necessidade comprovada, quando dever� ser produzido acordo expresso entre o AUXILIAR e a MANTENEDORA, � vedado, a esta, exigir, daquele, a realiza��o de trabalhos ou qualquer outra atividade aos domingos e feriados. Havendo o acordo e n�o sendo concedida folga compensat�ria, fica assegurada a remunera��o em dobro do trabalho realizado em tais dias, sem preju�zo do pagamento do repouso semanal remunerado. 

<tr><td class=titulo>10. ADICIONAL POR ATIVIDADES EM OUTROS MUNIC�PIOS 
<tr><td class=campo style="text-align:justify">Quando o AUXILIAR desenvolver suas atividades, em car�ter eventual, a servi�o da mesma MANTENEDORA, em munic�pio diferente daquele onde foi contratado e onde ocorre a presta��o habitual do trabalho, dever� receber um adicional de 25% (vinte e cinco por cento) sobre o total de sua remunera��o no novo munic�pio. Quando o AUXILIAR voltar a prestar servi�os no munic�pio de origem, cessar� a obriga��o do pagamento deste adicional. 
<br><b>Par�grafo primeiro</b> � Nos casos em que ocorrer a transfer�ncia definitiva do AUXILIAR, aceita livremente por este, em documento firmado entre as partes, n�o haver� a incid�ncia do adicional referido no �caput�, obrigando-se a MANTENEDORA a efetuar o pagamento de um �nico sal�rio mensal integral, ao AUXILIAR, no ato de transfer�ncia, a t�tulo de ajuda de custo. 
<br><b>Par�grafo segundo</b> � Fica assegurada a garantia de emprego pelo per�odo de 6 (seis) meses ao AUXILIAR transferido de munic�pio, contados a partir do in�cio do trabalho e/ou da efetiva��o da transfer�ncia. 
<br><b>Par�grafo terceiro</b> � Caso a MANTENEDORA desenvolva atividade acad�mica em munic�pios considerados conurbanados, poder� solicitar isen��o do pagamento do adicional determinado no caput, desde que encaminhe material comprobat�rio ao SEMESP, para an�lise e delibera��o do Foro Conciliat�rio para Solu��o de Conflitos Coletivos, previsto na presente Conven��o. 

<tr><td class=titulo>11. DESCONTO DE FALTAS 
<tr><td class=campo style="text-align:justify">Na ocorr�ncia de faltas n�o amparadas na legisla��o, a MANTENEDORA poder� descontar, no m�ximo, o n�mero de horas em que o AUXILIAR esteve ausente e o DSR proporcional a essas horas, desde que a MANTENEDORA n�o tenha implantado o sistema de Banco de Horas conforme o disposto em cl�usula pr�pria da presente Conven��o Coletiva de Trabalho. 
<br><b>Par�grafo �nico</b> � � da compet�ncia e integral responsabilidade da MANTENEDORA estabelecer mecanismos de controle de faltas e de pontualidade do AUXILIAR, conforme a legisla��o vigente. 

<tr><td class=titulo>12. ATESTADOS M�DICOS E ABONO DE FALTAS 
<tr><td class=campo style="text-align:justify">A MANTENEDORA � obrigada a aceitar atestados fornecidos por m�dicos ou dentistas conveniados ou credenciados pela entidade sindical profissional, SUS ou, ainda, por profissionais conveniados com a pr�pria MANTENEDORA. 
<br><b>Par�grafo �nico</b> � Tamb�m ser�o aceitos atestados que tenham sido convalidados pelas entidades sindicais de trabalhadores abrangidos por esta norma, pelos profissionais de sa�de de departamento m�dico ou odontol�gico pr�prio ou conveniados �s mesmas. 

<tr><td class=titulo>13. ANOTA��ES NA CARTEIRA DE TRABALHO 
<tr><td class=campo style="text-align:justify">A MANTENEDORA est� obrigada a promover, em quarenta e oito horas, �s anota��es nas Carteiras de Trabalho de seus AUXILIARES, ressalvados eventuais prazos mais amplos permitidos por lei. 
<br><b>Par�grafo �nico</b> � � obrigat�ria a anota��o na CTPS das mudan�as provocadas por ascens�o em plano de carreira. 

<tr><td class=titulo>14. MUDAN�A DE CARGO OU FUN��O 
<tr><td class=campo style="text-align:justify">O AUXILIAR n�o poder� ser transferido de um cargo ou fun��o para outro, salvo com seu consentimento expresso e por escrito, sob pena de nulidade da referida transfer�ncia. 

<tr><td class=titulo>15. ABONO DE FALTAS POR CASAMENTO OU LUTO 
<tr><td class=campo style="text-align:justify">N�o ser�o descontadas, no curso de nove dias corridos, as faltas do AUXILIAR, por motivo de gala ou luto, este em decorr�ncia de falecimento de pai, m�e, filho(a), c�njuge, companheiro(a) e dependente juridicamente reconhecido. 
<br><b>Par�grafo �nico</b> � Em caso de falecimento de irm�o(�), sogro(a) e neto(a) os abonos ficar�o reduzidos a tr�s dias. 

<tr><td class=titulo>16. BOLSAS DE ESTUDO 
<tr><td class=campo style="text-align:justify">Todo AUXILIAR tem direito a bolsas de estudo integrais, incluindo matr�cula, para si, c�njuge, filhos ou dependentes legais, ambos entendidos como aqueles reconhecidos pela legisla��o do Imposto de Renda ou aqueles que estejam sob a guarda judicial do AUXILIAR e vivam sob sua depend�ncia econ�mica, devidamente comprovada. Os filhos ou dependentes legais do AUXILIAR poder�o usufruir as bolsas de estudo integrais, sem qualquer �nus, desde que n�o tenham 25 (vinte e cinco) anos completos ou mais na data da efetiva��o da matr�cula no curso superior. As bolsas de estudo s�o v�lidas para cursos de gradua��o, p�s-gradua��o ou seq�enciais existentes e administrados pela MANTENEDORA localizado(s) no mesmo munic�pio onde o AUXILIAR trabalha, observado o disposto nesta cl�usula e par�grafos seguintes. 
<br><b>Par�grafo primeiro</b> � O direito �s bolsas de estudo passa a vigorar ao t�rmino do contrato de experi�ncia, cuja dura��o n�o pode exceder de 90 (noventa) dias, conforme par�grafo �nico do artigo 445 da CLT. 
<br><b>Par�grafo segundo</b> � A mantenedora est� obrigada a conceder duas bolsas de estudo, sendo que, nos cursos de gradua��o ou seq�enciais, n�o ser� poss�vel que o bolsista conclua mais de um curso nesta condi��o. 
<br><b>Par�grafo terceiro</b> � A utiliza��o do benef�cio previsto nesta cl�usula � transit�ria e n�o-habitual e, por isso, n�o possui car�ter remunerat�rio e nem se vincula, para nenhum efeito, ao sal�rio ou remunera��o percebida pelo AUXILIAR nos termos do inciso XIX, do par�grafo 9� do artigo 214 do Decreto 3048, de 06 de maio de 1999 e do par�grafo 2� do artigo 458 da CLT, com a reda��o dada pela Lei 10.243, de 19 de junho de 2001. 
<br><b>Par�grafo quarto</b> � As bolsas de estudo ser�o mantidas quando o AUXILIAR estiver licenciado para tratamento de sa�de ou em gozo de licen�a mediante anu�ncia da MANTENEDORA, excetuado o disposto na cl�usula da presente Conven��o que trata sobre a Licen�a sem Remunera��o. 
<br><b>Par�grafo quinto</b> � No caso de falecimento do AUXILIAR, os dependentes que j� se encontram estudando em estabelecimento de ensino superior da MANTENEDORA continuar�o a gozar das bolsas de estudo at� o final do curso, ressalvado o disposto no par�grafo 8� desta cl�usula. 
<br><b>Par�grafo sexto</b> � No caso de dispensa sem justa causa durante o per�odo letivo, ficam garantidas ao AUXILIAR, at� o final do per�odo letivo, as bolsas de estudo j� existentes. 
<br><b>Par�grafo s�timo</b> � As bolsas de estudo integrais em cursos de p�s-gradua��o ou especializa��o existentes e administrados pela MANTENEDORA s�o v�lidas exclusivamente para o AUXILIAR, respeitados os crit�rios de sele��o exigidos para ingresso nos mesmos e obedecer�o �s seguintes condi��es: 
<blockquote style="margin-top:0;margin-bottom:0">a) os cursos stricto sensu ou de especializa��o que fixem um n�mero m�ximo de alunos por turma, s�o limitadas em 30% (trinta por cento) do total de vagas oferecidas; 
<br>b) nos cursos de p�s-gradua��o lato sensu n�o haver� limites de vagas. 
</blockquote>
Caso a estrutura do curso torne necess�ria a limita��o do n�mero de alunos ser� observado o disposto na al�nea a) deste par�grafo. 
<br><b>Par�grafo oitavo</b> � Os bolsistas que forem reprovados no per�odo letivo perder�o o direito � bolsa de estudo, voltando a gozar do benef�cio quando lograrem aprova��o no referido per�odo. As disciplinas cursadas em regime de depend�ncia ser�o de total responsabilidade do bolsista, arcando o mesmo com o seu custo.
<br><b>Par�grafo nono</b> � Considera-se adquirido o direito daquele AUXILIAR que j� esteja usufruindo bolsas de estudo em n�mero superior ao definido nesta cl�usula. 

<tr><td class=titulo>17. IRREDUTIBILIDADE SALARIAL 
<tr><td class=campo style="text-align:justify">� proibida a redu��o da remunera��o mensal ou de carga hor�ria do AUXILIAR, exceto quando ocorrer iniciativa expressa do mesmo. Em qualquer hip�tese, � obrigat�ria a concord�ncia formal e rec�proca, firmada por escrito. 
<br><b>Par�grafo �nico</b> � N�o havendo concord�ncia rec�proca, a parte que deu origem � redu��o prevista nesta cl�usula arcar� com a responsabilidade da rescis�o contratual. 

<tr><td class=titulo>18. UNIFORMES 
<tr><td class=campo style="text-align:justify">A MANTENEDORA dever� fornecer gratuitamente dois uniformes por ano, quando o seu uso for exigido. 

<tr><td class=titulo>19. LICEN�A SEM REMUNERA��O 
<tr><td class=campo style="text-align:justify">O AUXILIAR, com mais de 5 (cinco) anos ininterruptos de servi�o no estabelecimento ensino superior da MANTENEDORA, ter� direito a licenciar-se, sem direito � remunera��o, por um per�odo m�ximo de dois anos, n�o sendo este per�odo de afastamento computado para contagem de tempo de servi�o ou para qualquer outro efeito, inclusive legal. 
<br><b>Par�grafo primeiro</b> � A licen�a ou sua prorroga��o dever�o ser comunicadas � MANTENEDORA com anteced�ncia m�nima de 90 (noventa) dias, devendo especificar as datas de in�cio e t�rmino do afastamento. A licen�a s� ter� in�cio a partir da data expressa no comunicado, mantendo-se, at� a�, todas as vantagens contratuais. A inten��o de retorno do AUXILIAR � atividade dever� ser comunicada � MANTENEDORA no m�nimo 60 (sessenta) dias antes do t�rmino do afastamento. 
<br><b>Par�grafo segundo</b> � O AUXILIAR que tenha ou exer�a cargo de confian�a dever�, junto com o comunicado de licen�a, solicitar seu desligamento do cargo a partir do in�cio da licen�a. 
<br><b>Par�grafo terceiro</b> � Considera-se demission�rio o AUXILIAR que, ao t�rmino do afastamento, n�o retornar �s atividades. 

<tr><td class=titulo>20. LICEN�A � AUXILIAR ADOTANTE 
<tr><td class=campo style="text-align:justify">Nos termos da Lei n� 10.421, de 15 de abril de 2.002, ser� garantida licen�a maternidade �s AUXILIARES que vierem a adotar ou obtiverem guarda judicial de crian�as. 

<tr><td class=titulo>21. LICEN�A PATERNIDADE 
<tr><td class=campo style="text-align:justify">A licen�a paternidade ter� a dura��o de 5 dias.

<tr><td class=titulo>22. GARANTIA DE EMPREGO � GESTANTE 
<tr><td class=campo style="text-align:justify">Fica garantido de emprego � AUXILIAR gestante desde o in�cio da gravidez at� sessenta dias ap�s o t�rmino do afastamento legal. Em caso de dispensa, o aviso pr�vio come�ar� a contar a partir do t�rmino do per�odo de estabilidade. 

<tr><td class=titulo>23. CRECHES 
<tr><td class=campo style="text-align:justify">� obrigat�ria a instala��o de local destinado � guarda de crian�as de at� doze meses, quando a unidade de ensino da MANTENEDORA mantiver contratadas, em jornada integral, pelo menos trinta funcion�rias. A manuten��o da creche poder� ser substitu�da pelo pagamento do reembolso-creche, nos termos da legisla��o em vigor (CF, 7�, XXV, Artigo 389, par�grafo 1� da CLT e Portaria MTb n� 3296 de 03.09.86), ou ainda, a celebra��o de conv�nio com uma entidade reconhecidamente id�nea. 

<tr><td class=titulo>24. GARANTIAS AO AUXILIAR EM VIAS DE APOSENTADORIA 
<tr><td class=campo style="text-align:justify">Fica assegurado ao AUXILIAR que, comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria por tempo de contribui��o ou da aposentadoria por idade, a garantia de emprego durante o per�odo que faltar at� a aquisi��o do direito. 
<br><b>Par�grafo primeiro</b> � A garantia de emprego � devida ao AUXILIAR que esteja contratado pela MANTENEDORA h� pelo menos tr�s anos. 
<br><b>Par�grafo segundo</b> � A comprova��o � MANTENEDORA dever� ser feita mediante a apresenta��o de documento que ateste o tempo de servi�o. Este documento dever� ser emitido pelo INSS ou por pessoa credenciada junto ao �rg�o previdenci�rio. Se o AUXILIAR depender de documenta��o para realiza��o da contagem, ter� um prazo de 30 (trinta) dias, a contar da data prevista ou marcada para homologa��o da rescis�o contratual. 
<br><b>Par�grafo terceiro</b> � O contrato de trabalho do AUXILIAR s� poder� ser rescindido por m�tuo acordo homologado pelo sindicato ou por pedido de demiss�o. 
<br><b>Par�grafo quarto</b> � Havendo acordo formal entre as partes, o AUXILIAR poder� exercer outra fun��o compat�vel, durante o per�odo em que estiver garantido pela estabilidade. 
<br><b>Par�grafo quinto</b> � O aviso pr�vio, em caso de demiss�o sem justa causa, integra o per�odo de estabilidade previsto nesta cl�usula. 
<br><b>Par�grafo sexto</b> � Enquanto n�o ocorrer a comprova��o da documenta��o prevista nesta cl�usula, o contrato de trabalho ficar� suspenso. Caso o AUXILIAR n�o apresente a documenta��o at� 30 (trinta) dias ap�s a data prevista para homologa��o da rescis�o, a demiss�o ocorrer� sem o pagamento de qualquer indeniza��o adicional. Ocorrendo a comprova��o da documenta��o, a rescis�o contratual ser� cancelada e o AUXILIAR ser� reintegrado.

<tr><td class=titulo>25. MULTA POR ATRASO NA HOMOLOGA��O DA RESCIS�O CONTRATUAL 
<tr><td class=campo style="text-align:justify">A MANTENEDORA deve homologar a rescis�o contratual at� o 20� dia ap�s o t�rmino do aviso pr�vio, quando trabalhado, ou trinta dias ap�s o desligamento, quando houver dispensa do cumprimento de aviso pr�vio. O atraso na homologa��o obrigar� a MANTENEDORA ao pagamento de multa, em favor do AUXILIAR, correspondente a um m�s de sua remunera��o. A partir do vig�simo dia de atraso, haver� ainda multa di�ria de 0,2% (dois d�cimos percentuais) do sal�rio mensal. A MANTENEDORA est� desobrigada de pagar a multa quando o atraso vier a ocorrer, comprovadamente, por motivos alheios � sua vontade. 
<br><b>Par�grafo �nico</b> � A entidade sindical profissional est� obrigada a fornecer comprovante de comparecimento sempre que a MANTENEDORA se apresentar para homologa��o das rescis�es contratuais e comprovar a convoca��o do AUXILIAR. 

<tr><td class=titulo>26. DEMISS�O POR JUSTA CAUSA 
<tr><td class=campo style="text-align:justify">Quando houver demiss�o por justa causa, nos termos do art. 482, da CLT, a MANTENEDORA est� obrigada a determinar na carta-aviso o motivo f�tico que deu origem � dispensa. Caso contr�rio, ficar� descaracterizada a justa causa. 

<tr><td class=titulo>27. READMISS�O DO AUXILIAR 
<tr><td class=campo style="text-align:justify">O AUXILIAR que for readmitido para a mesma fun��o at� 12 (doze) meses ap�s o seu desligamento ficar� desobrigado de firmar contrato de experi�ncia. 

<tr><td class=titulo>28. INDENIZA��O POR DISPENSA IMOTIVADA 
<tr><td class=campo style="text-align:justify">O AUXILIAR demitido sem justa causa ter� direito a uma indeniza��o, al�m do aviso pr�vio legal de trinta dias e das indeniza��es previstas nesta Conven��o, quando forem devidas, nas condi��es abaixo especificadas: 
<blockquote style="margin-top:0;margin-bottom:0">a) 03 (tr�s) dias para cada ano trabalhado na MANTENEDORA; 
<br>b) aviso pr�vio adicional de quinze dias, caso o AUXILIAR tenha, no m�nimo, cinq�enta anos de idade e que, � data do desligamento, conte com pelo menos um ano de servi�o na MANTENEDORA. 
</blockquote>
<b>Par�grafo primeiro</b> � N�o ter� direito a indeniza��o prevista na al�nea �a� o AUXILIAR que tiver recebido, durante pelo menos um ano, pagamento mensal de adicional por tempo de servi�o decorrente de plano de cargos e sal�rios ou de anu�nio, q�inq��nio ou equivalente, cujo valor corresponda a, no m�nimo, 1% (um por cento) do valor do sal�rio, por ano trabalhado. A MANTENEDORA dever� apresentar, no momento da homologa��o, documentos que comprovem o pagamento ao AUXILIAR do referido adicional por tempo de servi�o. 
<br><b>Par�grafo segundo</b> � N�o ter� direito � indeniza��o assegurada na al�nea �b� do caput, o AUXILIAR que, na data de admiss�o na MANTENEDORA, contar com mais de cinq�enta anos de idade.
<br><b>Par�grafo terceiro</b> � Essas indeniza��es n�o contar�o, para nenhum efeito, como tempo de servi�o. 

<tr><td class=titulo>29. ATESTADOS DE AFASTAMENTO E SAL�RIOS 
<tr><td class=campo style="text-align:justify">Sempre que solicitada, a MANTENEDORA dever� fornecer ao AUXILIARES atestado de afastamento e sal�rio (AAS) previsto na legisla��o vigente. 

<tr><td class=titulo>30. F�RIAS 
<tr><td class=campo style="text-align:justify">As f�rias dos AUXILIARES ser�o determinadas nos termos da legisla��o que rege a mat�ria, pela dire��o da MANTENEDORA, sendo admitida a compensa��o dos dias de f�rias concedidos antecipadamente, em per�odo nunca inferior a 10 (dez) dias e nem mais que 2 (duas) vezes por ano. 
<br><b>Par�grafo primeiro</b> � Fica assegurado aos AUXILIARES o pagamento, quando do in�cio de suas f�rias, do sal�rio correspondente �s mesmas e do abono previsto no inciso XVII, artigo 7�, da Constitui��o Federal, no prazo previsto pelo artigo 145 da CLT, independentemente de solicita��o pelos mesmos. 
<br><b>Par�grafo segundo</b> � As f�rias, individuais ou coletivas, n�o poder�o ter seu in�cio coincidindo com domingos, feriados, dia de compensa��o do repouso semanal remunerado ou s�bados, quando esses n�o forem dias normais de trabalho. 

<tr><td class=titulo>31. DELEGADO REPRESENTANTE 
<tr><td class=campo style="text-align:justify">Em cada unidade que tenha mais de 50 AUXILIARES, a MANTENEDORA assegurar� elei��o de um Delegado Representante, que ter� garantia de emprego e sal�rios a partir da inscri��o de sua candidatura at� seis meses ap�s o t�rmino de sua gest�o. 
<br><b>Par�grafo primeiro</b> � O mandato do Delegado Representante ser� de um ano. 
<br><b>Par�grafo segundo</b> � A elei��o do Delegado Representante ser� realizada pela entidade sindical na unidade de ensino da MANTENEDORA, por voto direto e secreto. � exigido quorum de 50% (cinq�enta por cento) mais um dos AUXILIARES da unidade de ensino da MANTENEDORA onde a elei��o ocorrer. 
<br><b>Par�grafo terceiro</b> � A entidade sindical comunicar� a elei��o � MANTENEDORA, com anteced�ncia m�nima de sete dias corridos. Nenhum candidato poder� ser demitido a partir da data da comunica��o at� o t�rmino da apura��o. 
<br><b>Par�grafo quarto</b> � � condi��o necess�ria que os candidatos tenham, � data da elei��o, pelo menos um ano de servi�o na MANTENEDORA. 

<tr><td class=titulo>32. QUADRO DE AVISOS 
<tr><td class=campo style="text-align:justify">A MANTENEDORA dever� colocar � disposi��o da entidade sindical da categoria profissional quadro de avisos, em local vis�vel, para fixa��o de comunicados de interesse da categoria, sendo proibida a divulga��o de mat�ria pol�tico-partid�ria ou ofensiva a quem quer que seja.

<tr><td class=titulo>33. ASSEMBL�IAS SINDICAIS 
<tr><td class=campo style="text-align:justify">Todo AUXILIAR ter� direito a abono de faltas para o comparecimento �s assembl�ias da categoria. 
<br><b>Par�grafo primeiro</b> � Na vig�ncia desta Conven��o, os abonos est�o limitados, a dois s�bados e mais dois dias �teis, quando a assembl�ia n�o for realizada no munic�pio em que o AUXILIAR trabalhe para a MANTENEDORA. Caso a Assembl�ia ocorra fora do munic�pio em que o AUXILIAR trabalhe para MANTENEDORA, os abonos est�o limitados, a dois s�bados e dois per�odos. As duas assembl�ias realizadas durante os dias �teis dever�o ocorrer em per�odos distintos. 
<br><b>Par�grafo segundo</b> � A entidade sindical dever� informar � MANTENEDORA, por escrito, com anteced�ncia m�nima de quinze dias corridos. Na comunica��o dever�o constar a data e o hor�rio da assembl�ia. 
<br><b>Par�grafo terceiro</b> � Os dirigentes sindicais n�o est�o sujeitos ao limite previsto no par�grafo primeiro desta cl�usula. As aus�ncias decorrentes do comparecimento �s assembl�ias de suas entidades ser�o abonadas mediante comunica��o formal � MANTENEDORA. 
<br><b>Par�grafo quarto</b> � A MANTENEDORA poder� exigir dos AUXILIARES e dos dirigentes sindicais atestado emitido pela entidade sindical profissional, que comprove o seu comparecimento � assembl�ia. 

<tr><td class=titulo>34. CONGRESSOS, SIMP�SIOS E EQUIVALENTES 
<tr><td class=campo style="text-align:justify">Os abonos de falta para comparecimento a congressos, simp�sios e equivalentes ser�o concedidos mediante aceita��o por parte da MANTENEDORA, que dever� formalizar por escrito a dispensa do AUXILIAR. 
<br><b>Par�grafo �nico</b> � A participa��o do AUXILIAR nos eventos descritos no �caput� n�o caracterizar� atividade extraordin�ria. 

<tr><td class=titulo>35. CONGRESSO DA ENTIDADE SINDICAL PROFISSIONAL 
<tr><td class=campo style="text-align:justify">Na vig�ncia desta Conven��o, a entidade sindical promover� um evento de natureza pol�tica ou pedag�gica (Congresso ou Jornada). A MANTENEDORA abonar� as aus�ncias de seus AUXILIARES que participarem do evento, nos seguintes limites: 
<blockquote style="margin-top:0;margin-bottom:0">a) no estabelecimento de ensino superior que tenha at� 49 AUXILIARES, ser� garantido, o abono a um AUXILIAR; 
<br>b) no estabelecimento de ensino superior que tenha entre 50 e 99 AUXILIARES, ser� garantido, o abono a dois AUXILIARES; 
<br>c) no estabelecimento de ensino superior que tenha mais de 100 AUXILIARES, ser� garantido, o abono a tr�s AUXILIARES. 
</blockquote>
Tais faltas, limitadas ao m�ximo de dois dias �teis al�m do s�bado, ser�o abonadas mediante a apresenta��o de atestado de comparecimento fornecido pela entidade sindical. O AUXILIAR dever� repor as horas que, porventura, sejam necess�rias para complementa��o da sua jornada de trabalho.

<tr><td class=titulo>36. RELA��O NOMINAL 
<tr><td class=campo style="text-align:justify">Obriga-se a MANTENEDORA a encaminhar para entidade representativa da categoria profissional, conforme Precedentes Normativos n� 41 e 111, do Tribunal Superior do Trabalho, no prazo m�ximo de trinta dias contados da data do recolhimento da Contribui��o Sindical, a rela��o nominal dos AUXILIARES que integram seu quadro de funcion�rios acompanhada do valor do sal�rio mensal e das guias das contribui��es sindical e assistencial. 

<tr><td class=titulo>37. FORO CONCILIAT�RIO PARA SOLU��O DE CONFLITOS COLETIVOS 
<tr><td class=campo style="text-align:justify">Fica mantida a exist�ncia do Foro Conciliat�rio para Solu��o de Conflitos Coletivos, que tem como objetivo procurar resolver: 
<blockquote style="margin-top:0;margin-bottom:0">I - diverg�ncias trabalhistas; 
<br>II - incapacidade econ�mico-financeira da MANTENEDORA, no cumprimento de reajuste salarial e/ou de cl�usulas previstas na presente conven��o coletiva; 
<br>III � altera��o no prazo de pagamento de sal�rios. 
</blockquote>
<b>Par�grafo primeiro</b> � Havendo dificuldade no cumprimento da cl�usula de reajuste salarial ou diminui��o nos percentuais de reajustes salariais estipulados nesta conven��o coletiva ou defini��o de outro crit�rio de reajuste salarial proposto pela MANTENEDORA, a solicita��o da realiza��o do Foro dever� ser formalizada por escrito e instru�da com a documenta��o pertinente ao pedido. 
<br><b>Par�grafo segundo</b> � Para efeito do que estabelece os incisos I, II e III deste artigo, a MANTENEDORA, ao solicitar o FORO, deve encaminhar os motivos do pedido de libera��o do cumprimento da cl�usula em quest�o, acompanhada da competente documenta��o comprobat�ria, para an�lise e decis�o. 
<br><b>Par�grafo terceiro</b> � O Foro ser� composto paritariamente, por tr�s representantes do SEMESP e do SAAESP. As reuni�es dever�o contar, tamb�m, com as partes em conflito que, se assim o desejarem, poder�o delegar representantes para substitu�-las e/ou serem assistidas por advogados, com poderes espec�ficos para adotarem, em nome da Institui��o, as decis�es julgadas convenientes e necess�rias. 
<br><b>Par�grafo quarto</b> � O SEMESP e o SAAESP dever�o indicar os seus representantes no Foro num prazo de trinta dias a contar da assinatura desta Conven��o. 
<br><b>Par�grafo quinto</b> � Cada sess�o do Foro ser� realizada no prazo m�ximo de quinze dias a contar da solicita��o formal e obrigat�ria de qualquer uma das entidades que o comp�em. A data, o local e o hor�rio ser�o decididos pelas entidades sindicais envolvidas. O n�o comparecimento de qualquer uma das partes acarretar� no encerramento imediato das negocia��es, bem como na aplica��o na multa estabelecida no Par�grafo nono desta cl�usula. 
<br><b>Par�grafo sexto</b> � Nenhuma das partes envolvidas ingressar� com a��o na Justi�a do Trabalho durante as negocia��es de entendimento. 
<br><b>Par�grafo s�timo</b> � Na aus�ncia de solu��o do conflito ou na hip�tese de n�o comparecimento de qualquer uma das partes, a comiss�o respons�vel pelo Foro fornecer� certid�o atestando o encerramento da negocia��o.
<br><b>Par�grafo oitavo</b> � Na hip�tese de sucesso das negocia��es, a crit�rio do Foro, a MANTENEDORA ficar� desobrigada de arcar com a multa prevista no par�grafo 9 � (nono) desta cl�usula. 
<br><b>Par�grafo nono</b> � As decis�es do Foro ter�o efic�cia legal entre as partes acordantes. O descumprimento das decis�es assumidas gerar� multa a ser estabelecida no Foro, independentemente daquelas j� estabelecidas nesta Conven��o. 
<br><b>Par�grafo dez</b> � A entidade sindical ou a MANTENEDORA que deixar de comparecer ao FORO, uma vez convocada, pagar� uma multa de R$ 1.000,00 (hum mil reais), que reverter� em favor da parte presente. 

<tr><td class=titulo>38. COMISS�O PERMANENTE DE NEGOCIA��O 
<tr><td class=campo style="text-align:justify">Fica mantida a Comiss�o Permanente de Negocia��o constitu�da de forma parit�ria, por tr�s (3) representantes das entidades sindicais profissionais e econ�mica, com o objetivo de: 
<blockquote style="margin-top:0;margin-bottom:0">a) fiscalizar o cumprimento das cl�usulas vigentes; 
<br>b) elucidar eventuais diverg�ncias de interpreta��o das cl�usulas desta Conven��o; 
<br>c) discutir quest�es n�o-contempladas na norma coletiva; 
<br>d) deliberar, no prazo m�ximo de trinta dias a contar da data da solicita��o protocolizada no SEMESP, sobre modifica��o de pagamento da assist�ncia m�dico-hospitalar, conforme os par�grafos 1� (primeiro) e 3� (terceiro) da cl�usula relativa � mat�ria, constante desta norma coletiva; 
<br>e) criar subs�dios para a Comiss�o de Tratativas Salariais 2008, atrav�s da elabora��o de documentos para a defini��o das fun��es/atividades e o regime de trabalho dos AUXILIARES. 
<br>f) criar crit�rios para a regionaliza��o das negocia��es salariais referentes a 2008, bem como definir crit�rios diferenciados para elabora��o do instrumento normativo destinado �s entidades mantenedoras de Universidades, Centros Universit�rios, Faculdades, Institutos Superiores de Educa��o e Centros de Educa��o Tecnol�gicas. 
</blockquote>
<b>Par�grafo primeiro</b> � As entidades sindicais componentes da Comiss�o Permanente de Negocia��o indicar�o seus representantes, no prazo m�ximo de trinta dias corridos, a contar da assinatura da presente Conven��o. 
<br><b>Par�grafo segundo</b> � A Comiss�o Permanente de Negocia��o dever� reunir-se mensalmente, em calend�rio elaborado de comum acordo entre as partes, alternadamente nas sedes das entidades sindicais que a comp�em. Nos casos dispostos na letra �d� do caput, dever� haver convoca��o espec�fica pela entidade sindical patronal. 
<br><b>Par�grafo terceiro</b> � O n�o comparecimento da entidade sindical, profissional ou econ�mica, nas reuni�es previstas no par�grafo 2� (segundo) da presente cl�usula, implicar� na multa de R$ 2.000,00 (dois mil reais) por reuni�o, a qual reverter� em benef�cio da entidade presente � mesma.

<tr><td class=titulo>39. ACORDOS INTERNOS 
<tr><td class=campo style="text-align:justify">Ficam assegurados os direitos mais favor�veis decorrentes de acordos internos ou de acordos coletivos de trabalho celebrados entre a MANTENEDORA e a entidade sindical profissional. 

<tr><td class=titulo>40. ASSIST�NCIA M�DICO-HOSPITALAR 
<tr><td class=campo style="text-align:justify">A MANTENEDORA est� obrigada a assegurar, �s suas expensas, assist�ncia m�dico-hospitalar a todos os seus AUXILIARES, sendo-lhe facultada a escolha por plano de sa�de, seguro-sa�de ou conv�nios com empresas prestadoras de servi�os m�dico-hospitalares. Poder�, ainda, prestar a referida assist�ncia diretamente em se tratando de institui��es que disponham de servi�os de sa�de e hospitais pr�prios ou conveniados. Qualquer que seja a op��o feita, a assist�ncia m�dico-hospitalar deve assegurar as condi��es e os requisitos m�nimos que seguem relacionados: 
<blockquote style="margin-top:0;margin-bottom:0">1. Abrang�ncia � A assist�ncia m�dico-hospitalar deve ser realizada no munic�pio onde funciona o estabelecimento de ensino superior ou onde vive o AUXILIAR, a crit�rio da MANTENEDORA. Em casos de emerg�ncia, dever� haver garantia de atendimento integral em qualquer localidade do Estado de S�o Paulo ou fixa��o, em contrato, de formas de reembolso. 
<br>2. Coberturas m�nimas: 
<blockquote style="margin-top:0;margin-bottom:0">2.1 Quarto para quatro pacientes, no m�ximo. 
2.2 Consultas. 
2.3 Prazo de interna��o de 365 dias por ano (comum e UTI/CTI) 
2.4 Parto, independentemente do estado grav�dico. 
2.5 Mol�stias infecto-contagiosas que exijam interna��o. 
2.6 Exames laboratoriais, ambulatoriais e hospitalares. 
</blockquote>
    3. Car�ncia � N�o haver� car�ncia na presta��o dos servi�os m�dicos e laboratoriais. 
<br>4. Auxiliar ingressante � N�o haver� car�ncia para o AUXILIAR ingressante, independentemente do m�s em que for contratado. 
<br>5. Pagamento � A assist�ncia m�dico-hospitalar ser� garantida nos termos desta Conven��o, cabendo ao AUXILIAR, para usufruir dos benef�cios da Lei n� 9656/98, o pagamento de 10% das mensalidades da referida assist�ncia, com teto limite de R$ 8,00 (oito reais) por m�s, respeitado o estabelecido no par�grafo 1� (primeiro) desta cl�usula. 
</blockquote>
    <b>Par�grafo primeiro</b> � Caso a assist�ncia m�dico-hospitalar vigente na Institui��o venha a sofrer reajuste em virtude de poss�veis modifica��es estabelecidas em legisla��o que abranja o segmento � Lei 9.656, de 03 de junho de 1998 e MP 2.097-39, de 26 de abril de 2001 - ou que vierem a ser estabelecidas em lei, ou por mudan�a de empresa prestadora de servi�o, a pedido do corpo t�cnico-administrativo da Institui��o ou por quebra de contrato, unilateralmente, por parte da atual empresa prestadora de servi�o, a MANTENEDORA continuar� a contribuir com o valor mensal vigente at� a data da modifica��o, devendo o AUXILIAR arcar com o valor excedente, que ser� descontado em folha e consignado no comprovante de pagamento, nos termos do art. 462, da CLT. 
<br><b>Par�grafo segundo</b> � Caso ocorra mudan�a de empresa prestadora de servi�o, por decis�o unilateral da MANTENEDORA, com conseq�ente reajuste no valor vigente, o AUXILIAR estar� isento do pagamento do valor excedente, cabendo � MANTENEDORA prover integralmente a assist�ncia m�dico-hospitalar, sem nenhum �nus para o AUXILIAR. 
<br><b>Par�grafo terceiro</b> � Para efeito do disposto no Par�grafo primeiro desta cl�usula, caber� � MANTENEDORA remeter a documenta��o comprobat�ria � Comiss�o Permanente de Negocia��o para a devida homologa��o. 
<br><b>Par�grafo quarto</b> � Fica obrigado o AUXILIAR a optar pela presta��o de assist�ncia m�dico-hospitalar em uma �nica Institui��o de ensino, quando mantiver mais de um v�nculo empregat�cio como AUXILIAR no mesmo munic�pio ou munic�pios conurbanos. � necess�rio que o AUXILIAR se manifeste por escrito, com anteced�ncia m�nima de vinte dias, para que a MANTENEDORA possa proceder � suspens�o dos servi�os. 
<br><b>Par�grafo quinto</b> � Mediante pagamento complementar e ades�o facultativa, conforme o plano de atendimento m�dico-hospitalar e devidamente documentado, o AUXILIAR poder� optar pela amplia��o dos servi�os de sa�de garantidos nesta Conven��o Coletiva ou estend�-los a seus dependentes. 

<tr><td class=titulo>41. SAL�RIO DO AUXILIAR ADMITIDO PARA SUBSTITUI��O 
<tr><td class=campo style="text-align:justify">Ao AUXILIAR admitido em substitui��o a outro desligado, qualquer que tenha sido o motivo do seu desligamento, ser� garantido, sempre, sal�rio inicial igual ao menor sal�rio na fun��o existente no estabelecimento, curso, grau ou n�vel de ensino, respeitado o Plano de Cargos e Sal�rios da MANTENEDORA, sem serem consideradas eventuais vantagens pessoais. 

<tr><td class=titulo>42. MENOR SAL�RIO DA CATEGORIA 
<tr><td class=campo style="text-align:justify">Fica assegurado, a partir de 1� (primeiro) de abril de 2007, nos termos do inciso V, artigo 7�, da Constitui��o Federal, um menor sal�rio da categoria equivalente a R$ 529,80 (quinhentos e vinte e nove reais e oitenta centavos centavos) por jornada integral de trabalho (44 horas semanais). 
<br>A partir de 1� (primeiro) de julho de 2007, nos termos do inciso V, artigo 7�, da Constitui��o Federal, ser� assegurado um menor sal�rio da categoria equivalente a R$ 532,35 (quinhentos e trinta e dois reais e trinta e cinco centavos) por jornada integral de trabalho (44 horas semanais). 

<tr><td class=titulo>43. ABONO DE PONTO AO ESTUDANTE 
<tr><td class=campo style="text-align:justify">Fica assegurado o abono de faltas ao AUXILIAR estudante para presta��o de exames escolares, condicionado � pr�via comunica��o � MANTENEDORA e comprova��o posterior. 

<tr><td class=titulo>44. PRORROGA��O DA JORNADA DO ESTUDANTE 
<tr><td class=campo style="text-align:justify">Fica permitida a prorroga��o da jornada de trabalho ao AUXILIAR estudante, ressalvadas as hip�teses de conflito com hor�rio de freq��ncia �s aulas.

<tr><td class=titulo>45. ESTABILIDADE PROVIS�RIA DO ALISTANDO 
<tr><td class=campo style="text-align:justify">� assegurada aos AUXILIARES em idade de presta��o do servi�o militar estabilidade provis�ria, desde o alistamento at� sessenta dias ap�s a baixa. 

<tr><td class=titulo>46. AUXILIAR AFASTADO POR DOEN�A 
<tr><td class=campo style="text-align:justify">Ao AUXILIAR afastado do servi�o por doen�a devidamente atestada pela Previd�ncia Social ou por m�dico ou dentista credenciado pela MANTENEDORA, ser� garantido o emprego ou o sal�rio, a partir da alta, por igual per�odo ao do afastamento, limitado a 60 (sessenta) dias al�m do aviso pr�vio. 

<tr><td class=titulo>47. REFEIT�RIOS 
<tr><td class=campo style="text-align:justify">A MANTENEDORA que contar com mais de 300 (trezentos) AUXILIARES no mesmo estabelecimento de ensino superior por ela mantido e n�o conceder vale-refei��o, obriga-se a manter refeit�rio. 
<br><b>Par�grafo �nico</b> � No estabelecimento de ensino superior da MANTENEDORA em que trabalhem menos de 300 (trezentos) AUXILIARES ser� obrigat�rio assegurar-lhes condi��es de conforto e higiene por ocasi�o das refei��es. 

<tr><td class=titulo>48. CESTA B�SICA 
<tr><td class=campo style="text-align:justify">Fica assegurada aos AUXILIARES que percebam, at� 5 (cinco) sal�rios m�nimos por m�s, em jornada de 36 (trinta e seis) horas semanais, ou percebam, em jornada inferior, remunera��o proporcionalmente igual ou inferior ao limite fixado nesta cl�usula, a concess�o de uma cesta b�sica mensal de 26 kg, composta, no m�nimo, dos seguintes produtos n�o perec�veis: 
<div align="center"><table width=350>
<tr><td class=campo>Arroz            </td><td class=campo>�leo                </td><td class=campo>Macarr�o </td></tr>
<tr><td class=campo>Feij�o           </td><td class=campo>Caf�                </td><td class=campo>Sal </td></tr>
<tr><td class=campo>Farinha de Trigo </td><td class=campo>Farinha de Mandioca </td><td class=campo>Farinha de Milho </td></tr>
<tr><td class=campo>A��car           </td><td class=campo>Biscoito            </td><td class=campo>Pur� de Tomate </td></tr>
<tr><td class=campo>Tempero          </td><td class=campo>Achocolatado        </td><td class=campo>Leite em P� </td></tr>
<tr><td class=campo>Fub�             </td><td class=campo>Sardinha em Lata    </td><td class=campo>Sop�o </td></tr>
</table></div>
<br><b>Par�grafo primeiro</b> � As MANTENEDORAS que j� concedem vale-refei��o, conforme o determinado pelo PAT, est�o desobrigadas do fornecimento de cesta b�sica. 
<br><b>Par�grafo segundo</b> � Fica assegurada a concess�o de cesta b�sica durante as f�rias, licen�a maternidade e licen�a doen�a, bem como ser� garantido ao AUXILIAR demitido sem justa causa, na vig�ncia da presente Conven��o, a cesta b�sica referente ao per�odo de aviso pr�vio, ainda que indenizado. 

<tr><td class=titulo>49. COMPENSA��O SEMANAL DA JORNADA DE TRABALHO 
<tr><td class=campo style="text-align:justify">Fica permitida a compensa��o semanal da jornada de trabalho, nos termos da legisla��o que rege a mat�ria e obedecido o seguinte crit�rio:
<blockquote style="margin-top:0;margin-bottom:0">a) mediante ci�ncia, atrav�s do calend�rio anual a ser publicado pela MANTENEDORA, os AUXILIARES ser�o dispensados do cumprimento de sua jornada de trabalho em dias ali previstos, compensando-se as horas n�o trabalhadas com horas de trabalho complementares. 

<tr><td class=titulo>50. BANCO DE HORAS 
<tr><td class=campo style="text-align:justify">Nos termos da Lei n� 9.601, de 21 de janeiro de 1998, fica celebrado o Banco de Horas entre os AUXILIARES e as MANTENEDORAS, conforme documento anexo � presente Conven��o Coletiva de Trabalho. 
<br><b>Par�grafo primeiro</b> � As MANTENEDORAS que desejarem implantar o Banco de Horas, conforme o disposto no caput, dever�o comunicar � entidade representativa da categoria profissional a implanta��o do mesmo, sob pena de n�o o fazendo n�o ter validade a aplicabilidade do Banco de Horas. 
<br><b>Par�grafo segundo</b> � Caso a MANTENEDORA queira fazer altera��es no Banco de Horas devido as suas peculiaridades, os crit�rios, detalhes, prazos e datas de implanta��o ser�o objeto de Acordo Coletivo de Trabalho espec�fico, firmado entre a MANTENEDORA e seus AUXILIARES, com a participa��o da entidade sindical representativa da categoria profissional, na forma da legisla��o em vigor. 

<tr><td class=titulo>51. AUTORIZA��O PARA DESCONTO EM FOLHA DE PAGAMENTO 
<tr><td class=campo style="text-align:justify">O desconto do AUXILIAR em folha de pagamento somente poder� ser realizado, mediante sua autoriza��o, nos termos dos artigos 462 e 545 da CLT, quando os valores forem destinados ao custeio de pr�mios de seguro, planos de sa�de, mensalidades associativas ou outras que constem da sua expressa autoriza��o, desde que n�o haja previs�o expressa de desconto na presente norma coletiva. 
<br><b>Par�grafo �nico</b> � Encontra-se na entidade sindical profissional, � disposi��o da MANTENEDORA, c�pia de autoriza��o do AUXILIAR para o desconto da mensalidade associativa. 

<tr><td class=titulo>52. ESTABILIDADE PARA PORTADORES DE DOEN�AS GRAVES 
<tr><td class=campo style="text-align:justify">Fica assegurada, at� alta m�dica, considerada como aptid�o ao trabalho, ou eventual concess�o de aposentadoria por invalidez, estabilidade no emprego aos AUXILIARES acometidos por doen�as graves ou incur�veis e aos AUXILIARES portadores do v�rus HIV que vierem a apresentar qualquer tipo de infec��o ou doen�a oportunista, resultante da patologia de base. 
<br><b>Par�grafo �nico</b> � S�o consideradas doen�as graves ou incur�veis, a tuberculose ativa, aliena��o mental, esclerose m�ltipla, neoplasia maligna, cegueira definitiva, hansen�ase, cardiopatia grave, doen�a de Parkinson, paralisia irrevers�vel e incapacitante, espondiloartrose anquilosante, nefropatia grave, estados do Mal de Paget (oste�te deformante) e contamina��o grave por radia��o.

<tr><td class=titulo>53. N�CLEO INTERSINDICAL DE CONCILIA��O TRABALHISTA 
<tr><td class=campo style="text-align:justify">Fica mantido o N�cleo Intersindical de Concilia��o Trabalhista (regulamento anexo) que funcionar� no sentido de buscar a composi��o de conflitos no �mbito das rela��es entre as partes representadas pelas entidades signat�rias desta Conven��o, nos termos previstos pelo artigo 625-C da Consolida��o das Leis do Trabalho, com a reda��o dada pela Lei 9.958, de 12 de janeiro de 2000. 

<tr><td class=titulo>54. GARANTIAS AO AUXILIAR COM SEQUELAS E READAPTA��O 
<tr><td class=campo style="text-align:justify">Ser� garantida ao AUXILIAR acidentado no trabalho ou acometido por doen�a profissional, a perman�ncia na MANTENEDORA em fun��o compat�vel com seu estado f�sico, sem preju�zo da remunera��o antes percebida, desde que ap�s o acidente ou comprova��o da aquisi��o de doen�a profissional apresente, cumulativamente, redu��o da capacidade laboral, atestada por �rg�o oficial e que se tenha tornado incapaz de exercer a fun��o que anteriormente desempenhava, obrigado, por�m, o AUXILIAR nessa situa��o a participar dos processos de readapta��o e reabilita��o profissionais. 
<br><b>Par�grafo �nico</b> � O per�odo de estabilidade do AUXILIAR que se encontra participando dos processos de readapta��o e reabilita��o profissionais ser� o previsto em lei. 

<tr><td class=titulo>55. COMPET�NCIA DAS ENTIDADES SINDICAIS SIGNAT�RIAS 
<tr><td class=campo style="text-align:justify">Fica estabelecida a legalidade das entidades sindicais signat�rias para promover, perante a Justi�a do Trabalho e o Foro em Geral, a��es pl�rimas em nome dos AUXILIARES em nome pr�prio, ou ainda, como parte interessada, em caso de descumprimento de qualquer cl�usula aven�ada ou determinada nesta norma coletiva. 

<tr><td class=titulo>56. PRIMEIROS SOCORROS 
<tr><td class=campo style="text-align:justify">A MANTENEDORA obriga-se a manter materiais de primeiros socorros nos locais de trabalho e providenciar, por sua conta, a remo��o do AUXILIAR acidentado/doente para o atendimento m�dico-hospitalar. 

<tr><td class=titulo>57. FLEXIBILIZA��O DA JORNADA DE TRABALHO 
<tr><td class=campo style="text-align:justify">Poder� ser flexibilizada a carga hor�ria entre jornadas do AUXILIAR, quando no exerc�cio concomitante de fun��o docente e atividade administrativa, n�o havendo assim pagamento de sal�rios nos intervalos, quando o AUXILIAR n�o tenha trabalhado nos mesmos. 

<tr><td class=titulo>58. CONTRIBUI��O ASSISTENCIAL PROFISSIONAL 
<tr><td class=campo style="text-align:justify">Obriga-se a MANTENEDORA a promover o desconto no exerc�cio de 2007, na folha de pagamento de seus AUXILIARES, sindicalizados e/ou filiados ou n�o, para recolhimento em favor do SAAE, entidade legalmente representativa da categoria dos AUXILIARES, na base territorial conferida pela respectiva carta sindical ou pelo inciso I, artigo 8� da Constitui��o Federal, em conta especial, da import�ncia correspondente ao percentual estabelecido ou ao que vier a ser estabelecido na Assembl�ia Geral da categoria. O recolhimento ser� realizado obrigatoriamente pela pr�pria MANTENEDORA, em guias pr�prias, acompanhadas das correspondentes rela��es nominais e valores devidos. As import�ncias destinam-se � cria��o, manuten��o e amplia��o dos servi�os assistenciais do SAAE, na conformidade das assembl�ias gerais. 
<br><b>Par�grafo primeiro</b> � Quando a MANTENEDORA deixar de efetuar o recolhimento das contribui��es estabelecidas nesta cl�usula mediante decis�o da referida Assembl�ia Geral, incorrer� na obrigatoriedade do pagamento de multa, cujo valor corresponder� a 5% (cinco por cento) do total da import�ncia a ser recolhida para o SAAE, acrescida da parcela correspondente � varia��o da TR ou de outro �ndice que vier a substitu�-la, a partir do dia seguinte ao vencimento, cabendo � MANTENEDORA a integral responsabilidade pela multa e demais comina��es, n�o podendo as mesmas, de forma alguma, incidir sobre os sal�rios dos AUXILIARES. 
<br><b>Par�grafo segundo</b> � Eventuais discord�ncias dos AUXILIARES, nos termos do Precedente Normativo n� 74 do TST e da ementa do STF, prolatada nos autos do recurso extraordin�rio n� 220-700-1, RS, em 06 de outubro de 1998 e publicada no DJ, edi��o de 13 de novembro de 1998 e do Ac�rd�o de STF, de 07/11/2000, dever�o ser comunicadas oficialmente pelo pr�prio AUXILIAR ao SAAE, no prazo de 10 dias antes da efetiva��o do primeiro pagamento, j� reajustado, com c�pia � MANTENEDORA, sob pena de perderem efic�cia. 
<br><b>Par�grafo terceiro</b> � O SAAE encaminhar� ao SEMESP, no prazo m�ximo de 30 (trinta) dias ap�s a assinatura da presente Conven��o, ata da assembl�ia geral que fixou a contribui��o, os respectivos valores e a �poca do desconto e do recolhimento, sob pena de n�o ocorrer o referido recolhimento. 
<br><b>Par�grafo quarto</b> � O desconto e o recolhimento da contribui��o assistencial, bem como os respectivos valores, foram decididos com base nos textos legais acima mencionados, em assembl�ia especificamente convocada e amplamente divulgada atrav�s de editais publicados em jornal de grande circula��o estadual e regional e devidamente realizada, nos termos do artigo 513, �e�, da Consolida��o das Leis do Trabalho, que estabelece, como prerrogativa das entidades sindicais �impor contribui��es a todos aqueles que participam das categorias econ�micas ou profissionais ou das profiss�es liberais representadas�. 
<br>Eventuais oposi��es ao desconto em quest�o dever�o ser apresentadas, individualmente, �s entidades sindicais, at� 10 (dez) dias antes da efetiva��o do primeiro pagamento, j� reajustado, sendo que manifesta��es fora do prazo estabelecido ser�o consideradas ineptas. 

<tr><td class=titulo>59. MULTA POR DESCUMPRIMENTO DA CONVEN��O 
<tr><td class=campo style="text-align:justify">O descumprimento de cada cl�usula desta Conven��o obrigar� a MANTENEDORA ao pagamento de multa correspondente a 5% (cinco por cento) do sal�rio do AUXILIAR, acrescida de juros e corre��o monet�ria, para cada parte prejudicada. 
<br><b>Par�grafo �nico</b> � A MANTENEDORA est� desobrigada de arcar com o valor previsto nesta cl�usula, caso o artigo da Conven��o j� estabele�a uma multa pelo n�o cumprimento da mesma.

<tr><td class=titulo>60. DISPOSI��ES TRANSIT�RIAS 
<tr><td class=campo style="text-align:justify">Cada signat�rio da presente Conven��o nomear� 3 (tr�s) representantes que formar� uma Comiss�o que dever�o se reunir mensalmente, durante seis meses, especificamente para estudar sobre Assist�ncia M�dico-Hospitalar, no que se refere a sua implementa��o por interm�dio do Sindicato dos Auxiliares, mediante crit�rios a serem definidos. 
<br><b>Par�grafo �nico</b> � A reuni�es ser�o realizadas toda terceira segunda-feira de cada m�s, sendo a primeira em maio de 2007. 

<tr><td class=campo style="text-align:justify">Por estarem justos e acertados, assinam a presente Conven��o Coletiva de Trabalho, a qual ser� depositada, para fins de arquivo, na Delegacia Regional do Trabalho e Emprego no Estado de S�o Paulo, nos termos do artigo 614, da Consolida��o das Leis do Trabalho, de modo a surtir, de imediato, os seus efeitos legais. 

<tr><td class=campo style="text-align:justify">S�o Paulo, 15 de maio de 2007. 
<br>
<pre>
<br>Hermes Ferreira Figueiredo                         Geraldo Mugayar 
<br>Presidente do SEMESP                               Presidente da FETEE 
<br>
<br>Augusto Cezar Casseb                               Celso Soares Nogueira 
<br>Presidente SEMESP Rio Preto                        Sind. dos Aux. de Adm. Escolar do ABC 
<br>
<br>Luiz Carlos Cust�dio                               F�tima Aparecida Marins Silva 
<br>Sind. dos Prof. e Aux. Administrativos             Sind. dos Aux. de Adm. Escolar de Bauru 
<br>de Ara�atuba e Regi�o 
<br>
<br>Moacir Pereira                                     Ronaldi Torelli 
<br>Sind. dos Prof. e Aux. de Adm. Escolar             Sind. dos Prof. e Trab. em Educa��o 
<br>de Bragan�a Paulista                               de Dracena e Regi�o 
<br>
<br>C�ssio Antonio da Silva Tenani                     C�ssio Antonio da Silva Tenani 
<br>Sind. dos Prof. e Aux. Administrativos             Sind. dos Prof. e Aux. Administrativos de Jales 
<br>de Fernand�polis 
<br>
<br>Ayrton Onofre da Silva                             Jos� Roberto Marques de Castro 
<br>Sind. dos Trab. em Estab. de Ensino de Lins        Sind. dos Trab. em Estab. de Ensino de Mar�lia
<br>
<br>Jos� Cl�udio Chaves                                Jo�o Manoel dos Santos 
<br>Sind. dos Aux. de Adm. Escolar de                  Sind. dos Aux. de Adm. Escolar de Piracicaba 
<br>Mogi das Cruzes
<br>
<br>Ademir Rodrigues                                   Rita Theresinha de Miranda Furquim    
<br>Sind. dos Trab. em Estab. de Ensino de             Sind. dos Prof. e Aux. de Adm. Escolar 
<br>Presidente Prudente                                de Ribeir�o Preto 
<br>
<br>M�rcio Campos                                      Maria Aparecida Maganha 
<br>Sind. dos Aux. de Adm. Escolar de Santos           Sind. dos Prof. e Aux. Administrativos de 
<br>                                                   S�o Jos� dos Campos 
<br>
<br>Valdecir Zampolla Caetano                          Cl�udio Figueroba Raimundo 
<br>Sind. dos Aux. de Adm. Escolar de                  Sind. dos Aux. de Adm. Escolar de Sorocaba 
<br>S�o Jos� do Rio Preto 
<br>
<br>Armando Raphael D�Avoglio 
<br>Sind. dos Prof. e Aux. Administrativos 
<br>de Votuporanga
</pre>
</table>

<DIV style="page-break-after:always"></DIV>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td class=titulo align="center">ANEXO 01
<tr><td class=titulo align="center">ACORDO COLETIVO DE TRABALHO PARA A INSTITUI��O DE BANCO DE HORAS
<tr><td class=campo style="text-align:justify"><b>Cl�usula Primeira</b> � Fica estabelecido entre as MANTENEDORAS, neste ato representadas pelo SEMESP � Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior do Estado de S�o Paulo e os AUXILIARES DE ADMINISTRA��O ESCOLAR, neste ato representado pelas ENTIDADES SINDICAIS PROFISSIONAIS, signat�rias da Conven��o Coletiva de Trabalho 2007 a cria��o do BANCO DE HORAS. 
<tr><td class=campo style="text-align:justify"><b>Cl�usula Segunda</b> � A partir de 01 de mar�o de 2007, fica institu�do para a categoria dos AUXILIARES de Administra��o Escolar, o Sistema de Banco de Horas, com base na Lei 9.601/98, que deu nova reda��o ao � 2� do artigo 59 da Consolida��o das Leis do Trabalho e a ele (art. 59) acrescentou o � 3�. 
<br><b>� 1�</b> Ser� formado um banco, proveniente das horas trabalhadas al�m da jornada normal di�ria, as quais ser�o compensadas nos termos do presente Acordo. 
<br><b>� 2�</b> A composi��o do Banco de Horas se dar� mediante o ac�mulo, apurado por meio de cart�o de ponto, de horas credoras ou devedoras. 
<br><b>� 3�</b> As horas excedentes, a que se refere o par�grafo 2�, estar�o limitadas a 2 (duas) horas di�rias e 10 (dez) horas semanais, as quais ser�o acumuladas para futura compensa��o. 
<br><b>� 4�</b> Ser� permitido um saldo negativo de, no m�ximo, 20 horas a serem compensadas, conforme estabelecido nos par�grafos 6� a 12�. 
<br><b>� 5�</b> As horas que ultrapassarem o limite estabelecido no par�grafo 3� desta cl�usula ser�o remuneradas como horas extras, em conformidade com o regulado em cl�usula pr�pria da Conven��o Coletiva de Trabalho 2007. 
<br><b>� 6�</b> A compensa��o n�o poder� ocorrer nas F�rias, Feriados e Descanso Semanal Remunerado. 
<br><b>� 7�</b> Sempre que houver interesse das partes em que haja a compensa��o, tal solicita��o se dar� com anteced�ncia m�nima de 48 (quarenta e oito) horas. 
<br><b>� 8�</b> A cada 120 (cento e vinte) dias ser�o realizados balan�os para apura��o do saldo de horas e planejamento da compensa��o, devendo tal saldo ser informado ao AUXILIAR. Havendo interesse entre as partes, o saldo existente poder� ser transferido, todo ou em parte, para o balan�o do per�odo seguinte. Poder�, ainda, o saldo apurado ser remunerado como hora extra, conforme o disposto na cl�usula n. � 09 da Conven��o Coletiva de Trabalho 2007/08. 
<br><b>� 9�</b> A apura��o e compensa��o de saldo negativo obedecer� ao mesmo crit�rio do par�grafo anterior. 
<br><b>� 10.</b> Os atrasos, sa�das e faltas por motivo justificado e n�o previsto na legisla��o ou na CCT 2007/08, poder�o ser compensados no Banco de Horas, limitando-se em uma ocorr�ncia por semana.
<br><b>� 11.</b> Os AUXILIARES contratados por prazo determinado, bem como aqueles que est�o em per�odo de experi�ncia, n�o poder�o valer-se do sistema de Banco de Horas. 
<br><b>� 12.</b> Nos casos de desligamento de AUXILIARES durante a vig�ncia deste Acordo, obrigar-se-� a MANTENEDORA a pagar o adicional de Horas Extras sobre as horas n�o compensadas, calculadas sobre o valor da remunera��o na data da rescis�o. Na exist�ncia de horas a compensar (saldo negativo), conforme previsto nos par�grafos 6� e 9�, estas ser�o descontadas das verbas rescis�rias. 
<br><b>� 13.</b> Qualquer diverg�ncia na aplica��o deste Acordo dever� ser resolvida atrav�s da convoca��o do Foro para Solu��o de Conflitos Coletivos, conforme Cl�usula espec�fica da Conven��o Coletiva de Trabalho. 
<br><b>� 14.</b> A renova��o, altera��o ou rescis�o deste Acordo depender� de acordo escrito dos representantes das partes, antes de expirado seu prazo de validade. 
<br><b>� 15.</b> O prazo de vig�ncia desta cl�usula � de 12 (doze) meses, encerrando-se em 28 de fevereiro de 2008. 
<tr><td class=campo style="text-align:justify">(Data e local de assinatura, com identifica��o dos signat�rios)
</table> 

<DIV style="page-break-after:always"></DIV>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td class=titulo align="center">ANEXO 02
<tr><td class=titulo align="center">INSTRUMENTO DE ADITAMENTO DA CONVEN��O COLETIVA DE TRABALHO
<tr><td class=campo style="text-align:justify">REGULAMENTO DO N�CLEO INTERSINDICAL DE CONCILIA��O TRABALHISTA 
<tr><td class=campo style="text-align:justify">Regulamento para funcionamento do N�cleo Intersindical de Concilia��o Trabalhista entre o Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de S�o Paulo - SEMESP e o Sindicato dos Auxiliares de Administra��o Escolar de S�o Paulo � SAAESP. 
<tr><td class=campo style="text-align:justify">Atrav�s do presente Instrumento de Aditamento, as partes d�o cumprimento ao que foi estipulado no par�grafo primeiro da cl�usula 53 da Conven��o Coletiva de Trabalho firmada entre os representantes das MANTENEDORAS e dos AUXILIARES DE ADMINISTRA��O ESCOLAR, implementando a cria��o do N�cleo Intersindical de Concilia��o Trabalhista previsto na Lei n� 9.958/2000, tudo nos termos das seguintes cl�usulas e condi��es que t�m como certas e ajustadas. 
<tr><td class=campo style="text-align:justify"><b>1.</b> Fica mantido o N�cleo Intersindical de Concilia��o Trabalhista entre o Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de S�o Paulo � SEMESP e o Sindicato dos Auxiliares de Administra��o escolar de S�o Paulo � SAAESP � previsto na cl�usula 54 da Conven��o Coletiva de Trabalho entre estas partes, bem como, no artigo 625-A da Consolida��o das Leis do Trabalho. 
<tr><td class=campo style="text-align:justify"><b>2.</b> O N�cleo aqui mencionado ir� funcionar nesta cidade, � Rua Machado Bittencourt, 317, 12� andar, Vila Clementino. 
<tr><td class=campo style="text-align:justify"><b>3.</b> Os trabalhos do N�cleo obedecer�o ao presente Regulamento, aprovado pelos convenentes. 
<tr><td class=campo style="text-align:justify"><b>4.</b> O N�cleo Intersindical de Concilia��o da Educa��o, doravante denominado simplesmente de Comiss�o, funcionar� nos termos previstos na Lei n.� 9.958/2000, com a finalidade de servir de instrumento para a r�pida solu��o dos conflitos individuais de trabalho, na �rea do ensino superior particular de S�o Paulo/SP. 
<tr><td class=campo style="text-align:justify"><b>5.</b> Para acionar os pr�stimos da Comiss�o, o interessado dever� protocolar na sede de funcionamento da Comiss�o, pedido de interven��o conciliat�ria, em cinco vias, sendo uma para arquivo na Comiss�o, outra para a notifica��o da parte contr�ria e as restantes para as Entidades Sindicais signat�rias. 
<tr><td class=campo style="text-align:justify"><b>6.</b> Tal pedido dever� expor de modo sint�tico os fatos e os fundamentos da quest�o, bem como, os valores pretendidos pelo interessado em raz�o de tal formula��o. 
<tr><td class=campo style="text-align:justify"><b>7.</b> O interessado poder� fazer-se representar por advogado na apresenta��o do pedido inicial, bem como, fazer-se acompanhar de tal profissional quando da sess�o de concilia��o, nesta oportunidade, o demandado dever� comparecer na pessoa de seu representante legal, ou por preposto, com poderes espec�ficos para transigir e firmar o termo de concilia��o. 
<tr><td class=campo style="text-align:justify"><b>8.</b> Recebido o pedido de interven��o conciliat�ria, a Comiss�o fixar� de imediato, data e hora para a sess�o de concilia��o, saindo intimado o interessado e notificando-se a parte contr�ria por escrito. Tal sess�o dever� realizar-se no m�ximo em dez dias, a contar da data do protocolo. 
<tr><td class=campo style="text-align:justify"><b>9.</b> A concilia��o praticada perante a Comiss�o, n�o poder� ser de car�ter gen�rico, somente sendo admiss�vel homologar transa��es sobre mat�ria constante do pedido inicial, conforme disposto na cl�usula 6� do presente instrumento. Ser� permitido aos interessados, inclusive, ressalvar expressamente que a transa��o n�o abrange alguma quest�o especificamente destacada. 
<tr><td class=campo style="text-align:justify"><b>10.</b> N�o poder� ser objeto de concilia��o a demanda que versar exclusivamente sobre FGTS, devendo ser entregue ao demandante, ou ao seu representante legal, no ato da distribui��o da demanda, termo de frustra��o expondo os motivos da recusa. 
<tr><td class=campo style="text-align:justify"><b>11.</b> Versando a demanda sobre outros itens al�m do FGTS, e havendo concilia��o entre as partes, dever� o respectivo termo conter ressalva quanto ao FGTS, ou determinar o dep�sito dos valores devidos (principal, multa e valores previstos na Lei Complementar n.� 110/2001) na conta vinculada do demandante, sob pena de nulidade da concilia��o neste t�pico. 
<tr><td class=campo style="text-align:justify"><b>12.</b> Aberta a sess�o conciliat�ria, os membros da Comiss�o explicar�o �s partes presentes qual a natureza das fun��es do �rg�o, bem como, tecer�o as pondera��es necess�rias � media��o para a solu��o negocial do conflito.
<tr><td class=campo style="text-align:justify"><b>13.</b> Obtida ou n�o a concilia��o entre as partes, ser� lavrado o termo respectivo para as finalidades previstas no par�grafo segundo do artigo 625-D ou no artigo 625-E da Lei 9958/2000. 
<tr><td class=campo style="text-align:justify"><b>14.</b> O N�cleo dever� intentar realizar a sess�o de concilia��o no prazo de 10 (dez) dias, a contar da provoca��o do interessado. N�o se ultimando a tentativa em tal prazo, ser� fornecida certid�o negativa ao interessado para os fins de direito. 
<tr><td class=campo style="text-align:justify"><b>15.</b> Os trabalhos do N�cleo ser�o desenvolvidos por conciliadores indicados pelas Entidades Sindicais signat�rias, em n�mero de 3 (tr�s) para cada parte convenente. Em cada sess�o realizada, os interessados ser�o sempre atendidos por, pelo menos, dois conciliadores, sendo um representante da Entidade Sindical patronal e outro da Entidade Sindical profissional.
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