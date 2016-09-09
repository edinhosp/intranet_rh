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
<title>Convenção Coletiva 2005 - Professores</title>
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
<tr><td class=titulo align="center">CONVENÇÃO COLETIVA DE TRABALHO 2005
<tr><td class=titulo align="center">SEMESP
<tr><td class=titulo align="center">PROFESSORES 
<tr><td class=campo style="text-align:justify">Entre as partes, de um lado, o Sindicato dos Professores de São José do Rio Preto - SINPRO São Paulo; SINPRO de OSASCO; SINPRO de Santos e Região; SINPRO de Jundiaí; SINPRO de Guarulhos; ee a Federação dos Professores do Estado de São Paulo – FEPESP, entidades com bases territoriais e representatividades fixadas nas respectivas Cartas Sindicais e no que estabelece o inciso I do artigo 8º da Constituição Federal e de outro, o Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de São Paulo – SEMESP e SEMESP São José do Rio Preto, com representatividade fixada em seus registros sindicais, ao final assinados por seus representantes legais, devidamente autorizados pelas competentes Assembléias Gerais das respectivas categorias, fica estabelecida, nos termos do artigo 611 e seguintes da Consolidação das Leis do Trabalho e do artigo 8º, inciso VI da Constituição Federal, a presente CONVENÇÃO COLETIVA DE TRABALHO:

<tr><td class=titulo>1. ABRANGÊNCIA
<tr><td class=campo style="text-align:justify">Esta Convenção abrange a categoria econômica dos estabelecimentos particulares de ensino superior no Estado de São Paulo, aqui designados como <b>MANTENEDORA</b> e a categoria profissional diferenciada dos professores, aqui designada simplesmente como PROFESSOR.
<br><b>Parágrafo único</b> – A categoria dos PROFESSORES abrange todos aqueles que exercem a atividade docente, independentemente da denominação sob a qual a função for exercida. Considera-se atividade docente a função de ministrar aulas.

<tr><td class=titulo>2. Duração
<tr><td class=campo style="text-align:justify">Esta Convenção Coletiva de Trabalho terá duração de dois anos, com vigência de 1º de março de 2005 a 28 de fevereiro de 2007.
<br><b>Parágrafo único</b> – As cláusulas poderão ser reexaminadas na próxima data-base em virtude de problemas surgidos na sua aplicação ou do surgimento de normas legais a elas pertinentes.

<tr><td class=titulo>3. Reajuste Salarial em 1º de maio de 2005
<tr><td class=campo style="text-align:justify">Sobre os salários devidos em 1º de junho de 2004 será aplicado, a partir de 1º de maio de 2005, um reajuste salarial de 7,66% (sete virgula sessenta e seis por cento), observado o estabelecido na cláusula 4ª da presente Convenção.
<br><b>Parágrafo primeiro</b> - Fica estabelecido que o salário de 1º de maio de 2005, reajustado pelo índice definido nesta cláusula, servirá como base de cálculo para a data base de 1º de março de 2006.
<br><b>Parágrafo segundo</b> – Eventuais diferenças salariais resultantes da aplicação da presente norma coletiva, até a data da sua assinatura, deverão ser pagas até o dia 15 de setembro de 2005, sem incidência de multa convencional.

<tr><td class=titulo>4. REAJUSTE SALARIAL em 1º de março de 2006
<tr><td class=campo style="text-align:justify">Em 1º de março de 2006, as MANTENEDORAS deverão aplicar sobre os salários devidos em 1º de maio de 2005, o percentual definido pela média aritmética dos índices inflacionários do período compreendido entre 1º de março de 2005 e 28 de fevereiro de 2006, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV).
<br><b>Parágrafo primeiro</b> – Se a média aritmética dos índices inflacionários definida no caput superar 9,99% (nove virgula noventa e nove por cento), as MANTENEDORAS deverão aplicar, em 1º de março de 2006, sobre os salários devidos em 1º de maio de 2005, o reajuste de 9,99% (nove virgula noventa e nove por cento). O SEMESP, o SINPRO e a FEPESP definirão, em processo de negociação salarial, até o prazo máximo de 30 de abril de 2006, a forma de pagamento da parcela excedente a 9,99%.
<br><b>Parágrafo segundo</b> – O SEMESP, o SINPRO e a FEPESP comprometem-se a divulgar, em comunicado conjunto, até 20 de março de 2006, o percentual de reajuste salarial calculado pela fórmula definida no caput, bem como a forma de pagamento da parcela excedente a 9,99%.
<br><b>Parágrafo terceiro</b> – A base de cálculo para a data-base de 1º de março de 2007 será constituída pelos salários devidos em 1º de maio de 2005, reajustados em 2006 pela média aritmética dos índices inflacionários do período compreendido entre 1º de março de 2005 e 28 de fevereiro de 2006, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV).

<tr><td class=titulo>5. COMPENSAÇÕES SALARIAIS
<tr><td class=campo style="text-align:justify">Para 2005 será permitida a compensação de eventuais antecipações salariais concedidas no período de vigência da Convenção de 2004. Relativamente à Convenção de 2006, será permitida a compensação de eventuais antecipações salariais concedidas no período de vigência da Convenção de 2005.
<br><b>Parágrafo único</b> – Excetuam-se em ambos os casos aquelas que decorrerem de promoções, transferências, ascensão em plano de carreira e aqueles reajustes concedidos com cláusula expressa de não compensação.

<tr><td class=titulo>6. SALÁRIO DO <b>PROFESSOR</b> INGRESSANTE NA MANTENEDORA
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> não poderá contratar nenhum <b>PROFESSOR</b> por salário inferior ao limite salarial mínimo dos PROFESSORES mais antigos que possuam o mesmo grau de qualificação ou titulação de quem está sendo contratado, respeitado o quadro de carreira da MANTENEDORA.
<br><b>Parágrafo único</b> – Ao <b>PROFESSOR</b> admitido após 1º de março de 2005 e após 1º de março de 2006, respectivamente, serão concedidos os mesmos percentuais de reajustes e aumentos salariais estabelecidos nesta norma coletiva.

<tr><td class=titulo>7. COMPROVANTE DE PAGAMENTO
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> deverá fornecer ao PROFESSOR, mensalmente, comprovante de pagamento, devendo estar discriminados: a)identificação da <b>MANTENEDORA</b> e do estabelecimento de ensino; b)a identificação do PROFESSOR; c)a denominação da categoria, se houver faixas salariais diferenciadas; d)o valor da hora-aula; e)a carga horária semanal; f)a hora-atividade; g)outros eventuais adicionais; h)o descanso semanal remunerado; i)as horas extras realizadas; j)o valor do recolhimento do FGTS; l)o desconto previdenciário; m)outros descontos.

<tr><td class=titulo>8. HORA-ATIVIDADE
<tr><td class=campo style="text-align:justify">Fica mantido o adicional de 5% (cinco por cento) a título de hora-atividade, destinado exclusivamente ao pagamento do tempo gasto pelo PROFESSOR, fora do estabelecimento de ensino, na preparação de aulas, provas e exercícios, bem como na correção dos mesmos.

<tr><td class=titulo>9. ADICIONAL NOTURNO
<tr><td class=campo style="text-align:justify">O trabalho noturno deve ser pago nas atividades realizadas após as 22 horas e corresponde a 25% (vinte e cinco por cento) do valor da hora-aula.

<tr><td class=titulo>10. HORAS EXTRAS
<tr><td class=campo style="text-align:justify">Considera-se atividade extra todo trabalho desenvolvido em horário diferente daquele habitualmente realizado na semana. As atividades extras devem ser pagas com adicional de 100% (cem por cento).
<br><b>Parágrafo primeiro</b> - Não é considerada atividade extra a participação em cursos de capacitação e aperfeiçoamento docente, desde que aceita livremente pelo PROFESSOR.
<br><b>Parágrafo segundo</b> - Serão pagas apenas como aulas normais, acrescidas do DSR e da hora-atividade, aquelas que forem adicionadas provisoriamente à carga horária habitual, decorrentes:
<blockquote style="margin-top:0;margin-bottom:0">a) da substituição temporária de um outro PROFESSOR, com duração predeterminada, decorrente de licença médica, maternidade ou para estudos. Nestes casos, a substituição deverá ser formalizada através de documento firmado entre a <b>MANTENEDORA</b> e o <b>PROFESSOR</b> que aceitar realizá-la;
<br>b) de substituições eventuais de faltas de <b>PROFESSOR</b> responsável, desde que aceitas livremente pelo <b>PROFESSOR</b> substituto;
<br>c) de reposição de eventuais faltas que foram descontadas dos salários nos meses em que ocorreram;
<br>d) da realização de cursos eventuais ou de curta duração, inclusive cursos de dependência, e aceitas livremente, mediante documento firmado entre o <b>PROFESSOR</b> convidado a ministrá-los e a MANTENEDORA;
<br>e) do comparecimento a reuniões didático-pedagógicas, de avaliação e de planejamento, quando realizadas fora de seu horário habitual de trabalho, desde que aceito livremente pelo PROFESSOR.
</blockquote>
<b>Parágrafo terceiro</b> – A participação em Comissões Internas e Externas da Unidade de Ensino da MANTENEDORA, desde que aceita livremente pelo <b>PROFESSOR</b> mediante documento firmado entre a <b>MANTENEDORA</b> e o PROFESSOR, será remunerada como aula ou hora normal, acrescida de DSR.

<tr><td class=titulo>11. JANELAS
<tr><td class=campo style="text-align:justify">Considera-se janela a aula vaga existente no horário do <b>PROFESSOR</b> entre duas outras aulas ministradas no mesmo turno. O pagamento da janela é obrigatório, devendo o <b>PROFESSOR</b> permanecer à disposição da <b>MANTENEDORA</b> neste período, ressalvada a aceitação pelo PROFESSOR, através de acordo formalizado entre as partes antes do início das aulas, quando as janelas não serão pagas.
<br><b>Parágrafo único</b> – Ocorrendo a hipótese da ressalva supra e caso o <b>PROFESSOR</b> seja solicitado esporadicamente a ministrar aulas ou a desenvolver qualquer outra atividade inerente ao magistério, no horário de janelas não-pagas, essas atividades serão remuneradas como aulas extras, com adicional de 100% (cem por cento).

<tr><td class=titulo>12. ADICIONAL POR ATIVIDADES EM OUTROS MUNICÍPIOS
<tr><td class=campo style="text-align:justify">Quando o <b>PROFESSOR</b> desenvolver suas atividades a serviço da mesma <b>MANTENEDORA</b> em município diferente daquele onde foi contratado e onde ocorre a prestação habitual do trabalho, deverá receber um adicional de 25% (vinte e cinco por cento) sobre o total de sua remuneração no novo município. Quando o <b>PROFESSOR</b> voltar a prestar serviços no município de origem, cessará a obrigação no pagamento do adicional.
<br><b>Parágrafo primeiro</b> – Nos casos em que ocorrer a transferência definitiva do PROFESSOR, aceita livremente por este, em documento firmado entre as partes, não haverá a incidência do adicional referido no caput, obrigando-se a <b>MANTENEDORA</b> a efetuar o pagamento de um único salário mensal integral, ao PROFESSOR, no ato da transferência, a título de ajuda de custo.
<br><b>Parágrafo segundo</b> – Fica assegurada a garantia de emprego pelo período de seis meses ao <b>PROFESSOR</b> transferido de município, contados a partir do início do trabalho e/ou da efetivação da transferência.
<br><b>Parágrafo terceiro</b> – Caso a <b>MANTENEDORA</b> desenvolva atividade acadêmica em municípios considerados conurbados, poderá solicitar isenção do pagamento do adicional determinado no caput, desde que encaminhe material comprobatório ao SEMESP e SEMESP – SÃO JOSÉ DO RIO PRETO, para análise e deliberação do Foro Conciliatório para Solução de Conflitos Coletivos, previsto na cláusula 46 da presente Convenção.

<tr><td class=titulo>13. COMPOSIÇÃO DO SALÁRIO MENSAL DO PROFESSOR
<tr><td class=campo style="text-align:justify">O salário do <b>PROFESSOR</b> é composto, no mínimo, por três itens: o salário base, o descanso semanal remunerado (DSR)e a hora-atividade.
<br>O salário base é calculado pela seguinte equação: número de aulas semanais multiplicado por 4,5 semanas e multiplicado, ainda, pelo valor da hora-aula (artigo 320, parágrafo 1º da CLT).
<br>O DSR corresponde a 1/6 (um sexto) do salário base, acrescido, quando houver, do total de horas extras e do adicional noturno (Lei 605/49).
<br>A hora-ativida decorresponde a 5% (cinco por cento) do total obtido com a somatória de todos os valores acima referidos.
<br><b>Parágrafo único</b> – A remuneração adicional do <b>PROFESSOR</b> pelo exercício concomitante de função não-docente obedecerá aos critérios estabelecidos entre a <b>MANTENEDORA</b> e o <b>PROFESSOR</b> que aceitar o cargo.

<tr><td class=titulo>14. Duração da Hora-aula
<tr><td class=campo style="text-align:justify">A duração da hora-aula poderá ser de, no máximo, cinqüenta minutos.
<br><b>Parágrafo primeiro</b> – Como exceção ao disposto no caput, a hora-aula poderá ter a duração de sessenta minutos nos cursos tecnológicos, desde que tenham sido autorizados ou reconhecidos com essa determinação expressa e cujos PROFESSORES desses cursos tenham sido contratados nessa condição.
<br><b>Parágrafo segundo</b> – As MANTENEDORAS de Instituições de Ensino que possuam cursos tecnológicos nas condições definidas no parágrafo 1º desta cláusula deverão apresentar à Comissão Permanente de Negociação definida na cláusula 47 da presente Convenção, até o dia 15 de agosto de 2005, a documentação de autorização ou reconhecimento do curso com a determinação expressa de hora-aula com duração de sessenta minutos sob pena de, em não o fazendo, estar sujeita à majoração do valor do salário-aula de acordo com o que estabelece o parágrafo 4º desta cláusula.
<br><b>Parágrafo terceiro</b> – Caso a Comissão Permanente de Acompanhamento delibere não ter havido determinação expressa do Ministério da Educação para que a duração da hora-aula dos cursos tecnológicos seja de sessenta minutos, a <b>MANTENEDORA</b> deverá majorar o salário-aula de acordo com o que estabelece o parágrafo 4º desta cláusula.
<br><b>Parágrafo quarto</b> – Em caso de ampliação da duração da hora-aula vigente, respeitado o limite previsto no caput desta cláusula, a <b>MANTENEDORA</b> deverá acrescer ao salário-aula já pago, valor proporcional ao acréscimo do trabalho.

<tr><td class=titulo>15. CARGA HORÁRIA
<tr><td class=campo style="text-align:justify">Quando a <b>MANTENEDORA</b> e o <b>PROFESSOR</b> contratarem carga diária de aulas superior aos limites previstos no artigo 318 da CLT, o excedente à carga horária legal será remunerado como aula normal, acrescido de DSR, hora-atividade e vantagens pessoais.
<br><b>Parágrafo único</b> - Poderá ser flexibilizada a carga horária do <b>PROFESSOR</b> entre jornadas, no exercício de sua função docente e concomitantemente com a atividade administrativa, não havendo assim pagamento, no intervalo, de horas aulas e salários, quando o <b>PROFESSOR</b> não tenha trabalhado no referido intervalo.

<tr><td class=titulo>16. PRAZO PARA PAGAMENTO DE SALÁRIOS
<tr><td class=campo style="text-align:justify">Os salários deverão ser pagos, no máximo, até o quinto dia útil do mês subseqüente ao trabalhado.
<br><b>Parágrafo único</b> – O não-pagamento dos salários no prazo obriga a <b>MANTENEDORA</b> a pagar multa diária, em favor do PROFESSOR, no valor de 1/50 (um cinqüenta avos) de seu salário mensal.

<tr><td class=titulo>17. DESCONTO DE FALTAS
<tr><td class=campo style="text-align:justify">Na ocorrência de faltas, a <b>MANTENEDORA</b> poderá descontar do salário do PROFESSOR, no máximo, o número de aulas em que o mesmo esteve ausente, o DSR (1/6), a hora-atividade e demais vantagens pessoais proporcionais a estas aulas.
<br><b>Parágrafo único</b> – É da competência e de integral responsabilidade da <b>MANTENEDORA</b> estabelecer mecanismos de controle de faltas e de pontualidade dos PROFESSORES, conforme a legislação vigente.

<tr><td class=titulo>18. ATESTADOS MÉDICOS E ABONO DE FALTAS
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> está obrigada a aceitar atestados fornecidos por médicos ou dentistas credenciados pelo SINPRO, SUS ou, ainda, profissionais conveniados com a própria MANTENEDORA.
<br><b>Parágrafo único</b> – Também serão aceitos atestados que tenham sido convalidados pelos profissionais de saúde do departamento médico ou odontológico do SINPRO ou conveniados a ele.

<tr><td class=titulo>19. ANOTAÇÕES NA CARTEIRA DE TRABALHO
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> está obrigada a promover, em quarenta e oito horas, as anotações nas Carteiras de Trabalho de seus PROFESSORES, ressalvados eventuais prazos mais amplos permitidos por lei.
<br><b>Parágrafo único</b> – É obrigatória a anotação na Carteira de Trabalho das mudanças provocadas por ascensão em plano de carreira ou alteração de titulação.

<tr><td class=titulo>20. MUDANÇA DE DISCIPLINA
<tr><td class=campo style="text-align:justify">O <b>PROFESSOR</b> não poderá ser transferido de uma disciplina para outra, salvo com seu consentimento expresso e por escrito, sob pena de nulidade da referida transferência.

<tr><td class=titulo>21. Redução de carga horária por extinção ou supressão de disciplina, classe ou turma
<tr><td class=campo style="text-align:justify">Ocorrendo supressão de disciplina, classe ou turma, em virtude de alteração na estrutura curricular prevista ou autorizada pela legislação vigente ou por dispositivo regimental devidamente aprovado por órgão colegiado da Instituição de Ensino, o <b>PROFESSOR</b> da disciplina classe ou turma deverá ser comunicado da redução da sua carga horária, por escrito, com antecedência mínima de 30 (trinta) dias do início do período letivo e terá prioridade para preenchimento de vaga existente em outra classe ou turma ou em outra disciplina para a qual possua habilitação legal.
<br><b>Parágrafo primeiro</b> – O <b>PROFESSOR</b> deverá manifestar por escrito, no prazo máximo de 5 (cinco) dias após a comunicação da MANTENEDORA, a não-aceitação da transferência de disciplina ou de classe ou turma ou da redução parcial de sua carga horária. A ausência de manifestação do <b>PROFESSOR</b> caracterizará a sua aceitação.
<br><b>Parágrafo segundo</b> – Caso o <b>PROFESSOR</b> não aceite a transferência para outra disciplina, classe ou turma ou a redução parcial de carga horária, a <b>MANTENEDORA</b> deverá manter a carga horária semanal existente ou, em caso contrário, proceder à rescisão do contrato de trabalho, por demissão sem justa causa.

<tr><td class=titulo>22. REDUÇÃO DE CARGA HORÁRIA POR DIMINUIÇÃO DE CARGA HORÁRIA POR DIMINUIÇÃO DO NÚMERO DE ALUNOS MATRICULADOS 
<tr><td class=campo style="text-align:justify">Na ocorrência de diminuição do número de alunos matriculados que venha a caracterizar a supressão de turmas, curso ou disciplina, o <b>PROFESSOR</b> do curso em questão deverá ser comunicado, por escrito, da redução parcial ou total de sua carga horária até o final da segunda semana de aulas do período letivo.
<br><b>Parágrafo primeiro</b> - O <b>PROFESSOR</b> deverá manifestar, também por escrito, a aceitação ou não da redução parcial de carga horária no prazo máximo de cinco dias após a comunicação da MANTENEDORA. A ausência de manifestação do <b>PROFESSOR</b> caracterizará a sua não-aceitação.
<br><b>Parágrafo segundo</b> - Caso o <b>PROFESSOR</b> aceite a redução parcial de carga horária, deverá formalizar documento junto à <b>MANTENEDORA</b> e, em não aceitando, a <b>MANTENEDORA</b> deverá proceder à rescisão do contrato de trabalho, por demissão sem justa causa, caso seja mantida a redução parcial de carga horária.
<br><b>Parágrafo terceiro</b> - Na hipótese de rescisão contratual, por demissão sem justa causa, o aviso prévio será indenizado, estando a <b>MANTENEDORA</b> desobrigada do pagamento do disposto na cláusula 29 da presente Convenção - Garantia Semestral deSalários.
<br><b>Parágrafo quarto</b> - Não ocorrendo redução do número de alunos matriculados que venha a caracterizar supressão do curso, de turma ou de disciplina, a <b>MANTENEDORA</b> que reduzir a carga horária do <b>PROFESSOR</b> estará sujeita ao disposto na cláusula 29 - Garantia Semestral de Salários - quando ocorrer a rescisão do contrato de trabalho do PROFESSOR.

<tr><td class=titulo>23. ABONO DE FALTAS POR CASAMENTO OU LUTO
<tr><td class=campo style="text-align:justify">Não serão descontadas, no curso de nove dias corridos, as faltas do PROFESSOR, por motivo de gala ou luto, este em decorrência de falecimento de pai, mãe, filho, cônjuge, companheira (o) e dependente juridicamente reconhecido.

<tr><td class=titulo>24. IRREDUTIBILIDADE SALARIAL
<tr><td class=campo style="text-align:justify">É proibida a redução de remuneração mensal ou de carga horária, ressalvada a ocorrência do disposto nas cláusulas 21 e 22 da presente Convenção, ou ainda, quando ocorrer iniciativa expressa do PROFESSOR. Em qualquer hipótese, é obrigatória a concordância recíproca, firmada por escrito.
<br><b>Parágrafo primeiro</b> - Não havendo concordância recíproca, a parte que deu origem à redução prevista nesta cláusula arcará com a responsabilidade da rescisão contratual.
<br><b>Parágrafo segundo</b> – Outras atividades, ainda que inerentes ao trabalho docente, que não sejam as de ministrar aulas, de duração temporária e determinada, poderão ser regulamentadas por contrato entre as partes, contendo a caracterização da atividade, o início e a previsão do término. 

<tr><td class=titulo>25. UNIFORMES
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> deverá fornecer gratuitamente dois uniformes por ano, quando o seu uso for exigido.

<tr><td class=titulo>26. LICENÇA SEM REMUNERAÇÃO
<tr><td class=campo style="text-align:justify">O <b>PROFESSOR</b> com mais de cinco anos ininterruptos de serviço na <b>MANTENEDORA</b> terá direito a licenciar-se, sem direito à remuneração, por um período máximo de dois anos, não sendo este período de afastamento computado para contagem de tempo de serviço ou para qualquer outro efeito, inclusive legal.
<br><b>Parágrafo primeiro</b> – A licença ou sua prorrogação deverá ser comunicada por escrito, à MANTENEDORA, com antecedência mínima de noventa dias do período letivo, devendo especificar as datas de início e término do afastamento. A licença só terá início a partir da data expressa no comunicado, mantendo-se, até aí, todas as vantagens contratuais. A intenção de retorno do <b>PROFESSOR</b> à atividade deverá ser comunicada à MANTENEDORA, no mínimo, sessenta dias antes do término do afastamento.
<br><b>Parágrafo segundo</b> – O término do afastamento deverá coincidir com o início do período letivo.
<br><b>Parágrafo terceiro</b> – O <b>PROFESSOR</b> que tenha ou exerça cargo de confiança deverá, junto com o comunicado de licença, solicitar seu desligamento do cargo a partir do início do período de licença.
<br><b>Parágrafo quarto</b> – Considera-se demissionário o <b>PROFESSOR</b> que, ao término do afastamento, não retornar às atividades docentes.
<br><b>Parágrafo quinto</b> – Ocorrendo a dispensa sem justa causa ao término da licença, o <b>PROFESSOR</b> não terá direito à Garantia Semestral de Salários, prevista na cláusula 29 da presente Convenção.

<tr><td class=titulo>27. LICENÇA À PROFESSORA ADOTANTE
<tr><td class=campo style="text-align:justify">Nos termos da Lei 10.421, de 15 de abril de 2002, será assegurada licença maternidade à professora que vier a adotar ou obtiver guarda judicial de crianças, garantido o emprego no período em que a licença for concedida.

<tr><td class=titulo>28. LICENÇA PATERNIDADE
<tr><td class=campo style="text-align:justify">A licença paternidade terá duração de cinco dias.

<tr><td class=titulo>29. GARANTIA SEMESTRAL DE SALÁRIOS
<tr><td class=campo style="text-align:justify">Ao <b>PROFESSOR</b> demitido sem justa causa, a <b>MANTENEDORA</b> garantirá:
<blockquote style="margin-top:0;margin-bottom:0">
•no primeiro semestre, a partir de 1º de janeiro, os salários integrais até o dia 30 de junho; 
<br>•no segundo semestre, os salários integrais até o dia 31 de dezembro, ressalvado o parágrafo 4º. 
</blockquote>
<b>Parágrafo primeiro</b> - Não terá direito à Garantia Semestral de Salários o <b>PROFESSOR</b> que, na data da comunicação da dispensa, contar com menos de 18 (dezoito) meses de serviço prestado à MANTENEDORA, ressalvado o parágrafo 4º desta cláusula.
<br><b>Parágrafo segundo</b> – No caso de demissões efetuadas no final do primeiro semestre letivo, para não ficar obrigada a pagar ao <b>PROFESSOR</b> os salários do segundo semestre a <b>MANTENEDORA</b> deverá observar as seguintes disposições:
<blockquote style="margin-top:0;margin-bottom:0">
•com aviso prévio a ser trabalhado, a demissão deverá ser formalizada com antecedência mínima de trinta dias do início das férias; 
<br>•sendo o aviso prévio indenizado, a demissão deverá ser formalizada até um dia antes do início das férias, ainda que as férias tenham seu início programado para o mês de julho, obedecendo ao que dispõe a cláusula 38 da presente Convenção. 
<br>•Os dias de aviso prévio que forem indenizados não contarão como tempo de serviço para efeito do pagamento da Garantia Semestral de Salários, conforme o estabelecido nesta cláusula.
</blockquote>
<b>Parágrafo terceiro</b> - No caso de demissões efetuadas no final do ano letivo, para não ficar obrigada a pagar ao <b>PROFESSOR</b> os salários do primeiro semestre do ano seguinte a <b>MANTENEDORA</b> deverá observar as seguintes disposições:
<blockquote style="margin-top:0;margin-bottom:0">
•com aviso prévio a ser trabalhado, a demissão deverá ser formalizada com antecedência mínima de trinta dias do início do recesso escolar; 
<br>•sendo o aviso prévio indenizado, a demissão deverá ser formalizada até um dia antes do início do recesso escolar. 
<br>•Os dias de aviso prévio que forem indenizados não contarão como tempo de serviço para efeito do pagamento da Garantia Semestral de Salários, conforme o estabelecido nesta cláusula.
</blockquote>
<b>Parágrafo quarto</b> - Quando as demissões ocorrerem a partir de 16 de outubro, a <b>MANTENEDORA</b> pagará, independentemente do tempo de serviço do professor, valor correspondente à remuneração devida até o dia 18 de janeiro do ano subseqüente, inclusive, ressalvados os contratos de experiência e por prazo determinado, estes últimos válidos somente nos casos de substituição temporária, conforme o disposto na alínea a) do parágrafo 2º da cláusula 10ª da presente Convenção.
<br><b>Parágrafo quinto</b> - Os salários complementares previstos nesta cláusula terão natureza indenizatória, não integrando, para nenhum efeito legal, o tempo de serviço do professor.
<br><b>Parágrafo sexto</b> - O aviso prévio de trinta dias previsto no artigo 487 da CLT já está integrado às indenizações tratadas nesta cláusula.

<tr><td class=titulo>30. GARANTIA DE EMPREGO À GESTANTE
<tr><td class=campo style="text-align:justify">É proibida a dispensa arbitrária ou sem justa causa da PROFESSORA gestante, desde o início da gravidez até sessenta dias após o término do afastamento legal.
O aviso prévio começará a contar a partir do término do período de estabilidade.

<tr><td class=titulo>31. CRECHES
<tr><td class=campo style="text-align:justify">É obrigatória a instalação de local destinado à guarda de crianças de até seis meses, quando a <b>MANTENEDORA</b> mantiver contratadas, em jornada integral, pelo menos trinta funcionárias com idade superior a 16 anos. A manutenção da creche poderá ser substituída pelo pagamento do reembolso-creche, nos termos da legislação em vigor (artigo 389, parágrafo 1º da CLT e Portarias MTb nº 3296 de 03.09.86 e nº 670, de 27/08/97), ou ainda, a celebração de convênio com uma entidade reconhecidamente idônea.

<tr><td class=titulo>32. GARANTIAS AO <b>PROFESSOR</b> EM VIAS DE APOSENTADORIA
<tr><td class=campo style="text-align:justify">Fica assegurado ao <b>PROFESSOR</b> que, comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria integral por tempo de serviço ou da aposentadoria por idade, a garantia de emprego durante o período que faltar até a aquisição do direito.
<br><b>Parágrafo primeiro</b> – A garantia de emprego é devida ao <b>PROFESSOR</b> que esteja contratado pela <b>MANTENEDORA</b> há pelo menos três anos.
<br><b>Parágrafo segundo</b> – A comprovação à <b>MANTENEDORA</b> deverá ser feita mediante a apresentação de documento que ateste o tempo de serviço. Este documento deverá ser emitido pela Previdência Social ou por pessoa credenciada junto ao órgão previdenciário. Se o <b>PROFESSOR</b> depender de documentação para realização da contagem, terá um prazo de quarenta e cinco dias, a contar da data da comunicação da dispensa. Comprovada a solicitação de tal documentação, os prazos serão prorrogados até que a mesma seja emitida.
<br><b>Parágrafo terceiro</b> – O contrato de trabalho do <b>PROFESSOR</b> só poderá ser rescindido por mútuo acordo homologado pelo SINPRO ou pedido de demissão.
<br><b>Parágrafo quarto</b> – Havendo acordo formal entre as partes, o <b>PROFESSOR</b> poderá exercer outra função, inerente ao magistério, durante o período em que estiver garantido pela estabilidade.
<br><b>Parágrafo quinto</b> – O aviso prévio, em caso de demissão sem justa causa, integra o período de estabilidade previsto nesta cláusula.

<tr><td class=titulo>33. MULTA POR ATRASO NA HOMOLOGAÇÃO
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> deve homologar a rescisão contratual no dia seguinte ao término do aviso prévio, quando trabalhado, ou dez dias após o desligamento, quando houver dispensa do cumprimento de aviso prévio. O atraso na homologação obrigará a <b>MANTENEDORA</b> ao pagamento de multa, em favor do PROFESSOR, correspondente a um mês de sua remuneração, conforme o disposto no parágrafo 8º do artigo 477 da CLT. A partir do vigésimo dia de atraso, haverá ainda multa diária de 0,2% (dois décimos percentuais) do salário mensal. A <b>MANTENEDORA</b> está desobrigada de pagar a multa quando o atraso vier a ocorrer, comprovadamente, por motivos alheios a sua vontade.
<br><b>Parágrafo único</b> – O SINPRO está obrigado a fornecer comprovante de comparecimento sempre que a <b>MANTENEDORA</b> se apresentar para homologação das rescisões contratuais e comprovar a convocação do PROFESSOR.

<tr><td class=titulo>34. DEMISSÃO POR JUSTA CAUSA
<tr><td class=campo style="text-align:justify">Quando houver demissão por justa causa, nos termos do art. 482 da CLT, a <b>MANTENEDORA</b> está obrigada a determinar na carta-aviso o motivo que deu origem à dispensa. Caso contrário, fica descaracterizada a justa causa.

<tr><td class=titulo>35. READMISSÃO DO PROFESSOR
<tr><td class=campo style="text-align:justify">O <b>PROFESSOR</b> que for readmitido até doze meses após o seu desligamento ficará desobrigado de firmar contrato de experiência.

<tr><td class=titulo>36. INDENIZAÇÕES POR DISPENSA IMOTIVADA
<tr><td class=campo style="text-align:justify">O <b>PROFESSOR</b> demitido sem justa causa terá direito a uma indenização, além do aviso prévio legal de trinta dias e das indenizações previstas na cláusula 28 desta Convenção, quando forem devidas, nas condições abaixo especificadas:
<blockquote style="margin-top:0;margin-bottom:0">
a) três (03) dias para cada ano trabalhado na MANTENEDORA;
<br>b) aviso prévio adicional de quinze dias, caso o <b>PROFESSOR</b> tenha, no mínimo, cinqüenta anos de idade e que, à data do desligamento, conte com pelo menos um ano de serviço na MANTENEDORA.
</blockquote>
<b>Parágrafo primeiro</b> – Não estará obrigada ao pagamento da indenização prevista na alínea a) a <b>MANTENEDORA</b> que tiver garantido ao <b>PROFESSOR</b> demitido, durante pelo menos um ano, pagamento mensal de adicional por tempo de serviço decorrente de plano de cargos e salários ou de anuênio, qüinqüênio ou equivalente, cujo valor corresponda a, no mínimo, 1% (um por cento) do valor da hora-aula por ano trabalhado e, por conseqüência, do salário mensal.
<br><b>Parágrafo segundo</b> – Não terá direito à indenização assegurada na alínea b) do caput, o <b>PROFESSOR</b> que, na data de admissão na MANTENEDORA, contar com mais de cinqüenta anos de idade.
<br><b>Parágrafo terceiro</b> – Para fazer jus à isenção prevista no parágrafo primeiro desta cláusula, a <b>MANTENEDORA</b> deverá encaminhar à Comissão Permanente de Negociação definida na cláusula 46 desta Convenção, no prazo máximo de noventa dias a contar da data da assinatura da presente Convenção, documentação que comprove o plano de pagamento de adicional por tempo de serviço nas condições estabelecidas no referido parágrafo.
<br><b>Parágrafo quarto</b> – Para a <b>MANTENEDORA</b> que não estiver enquadrada nos parágrafos primeiro e segundo, o pagamento das verbas indenizatórias previstas nesta cláusula não será cumulativo, cabendo ao PROFESSOR, no desligamento, o maior valor monetário entre os previstos nas alíneas a) e b) do caput.
<br><b>Parágrafo quinto</b> – Essas indenizações não contarão, para nenhum efeito, como tempo de serviço.

<tr><td class=titulo>37. ATESTADOS DE AFASTAMENTO E SALÁRIOS
<tr><td class=campo style="text-align:justify">Sempre que solicitada, a <b>MANTENEDORA</b> deverá fornecer ao <b>PROFESSOR</b> atestado de afastamento e salário (AAS), previsto na legislação previdenciária.

<tr><td class=titulo>38. FÉRIAS
<tr><td class=campo style="text-align:justify">As férias anuais dos PROFESSORES serão coletivas, com duração de trinta dias corridos e gozadas em julho de 2005 e julho de 2006. Qualquer alteração deverá ser aprovada por órgão competente, conforme o estabelecido em Estatuto ou Regimento e deverá constar do calendário escolar.
<br><b>Parágrafo primeiro</b> – A <b>MANTENEDORA</b> está obrigada a pagar o salário das férias e o abono constitucional de 1/3 (um terço) até quarenta e oito horas antes do início das férias.
<br><b>Parágrafo segundo</b> – As férias não poderão ser iniciadas aos domingos, feriados, dias de compensação do descanso semanal remunerado e nem aos sábados, quando estes não forem dias normais de aula.

<tr><td class=titulo>39. RECESSO ESCOLAR
<tr><td class=campo style="text-align:justify">O recesso escolar anual é obrigatório e tem duração de trinta dias corridos. Na vigência da presente Convenção os recessos escolares serão gozados preferencialmente no mês de janeiro de 2006 e no mês de janeiro de 2007.
Durante o recesso escolar anual que não pode, de maneira alguma, coincidir com o período definido para as férias coletivas do ano respectivo, o <b>PROFESSOR</b> não poderá ser convocado para nenhum trabalho.
<br><b>Parágrafo primeiro</b> –Na vigência da presente Convenção, as instituições cujos calendários escolares, determinados pelo órgão competente conforme o estabelecido em Estatuto ou Regimento, não observarem o determinado pelo caput para o recesso escolar anual dos PROFESSORES, poderão concedê-lo em um período de, no mínimo vinte dias corridos e em, no máximo, mais dois períodos com igual número de dias corridos, desde que observem as seguintes condições:
<blockquote style="margin-top:0;margin-bottom:0">
a) Vinte dias corridos em janeiro de 2006 e os dois períodos com igual número de dias corridos, obrigatoriamente no período compreendido entre março de 2005 e fevereiro de 2006.
<br>b) Vinte dias corridos em janeiro de 2007 e os dois períodos com igual número de dias corridos, obrigatoriamente no período compreendido entre março de 2006 e fevereiro de 2007.
</blockquote>
<b>Parágrafo segundo</b> – No caso dos calendários escolares preverem a divisão do recesso escolar dos PROFESSORES, os períodos definidos na conformidade do parágrafo primeiro não poderão ser iniciados aos domingos, feriados, dias de compensação do descanso semanal remunerado e nem aos sábados, quando estes não forem dias normais de aulas.
<br><b>Parágrafo terceiro</b> – As Instituições cujas atividades não possam ser interrompidas, tais como aquelas desenvolvidas em hospital, clínica, laboratório de análise, escritórios experimentais, pesquisas, dentre outros, ou que ministrem cursos em que sejam utilizadas instalações específicas ou que prestem atendimento à comunidade que não pode ser suspenso, poderão conceder aos PROFESSORES o recesso escolar anual definido no caput de maneira escalonada ao longo de cada ano.
<br><b>Parágrafo quarto</b> – Os calendários escolares que definirão os períodos de recesso escolar dos PROFESSORES serão obrigatoriamente divulgados aos PROFESSORES até o início de cada período letivo.

<tr><td class=titulo>40. DELEGADO REPRESENTANTE
<tr><td class=campo style="text-align:justify">Em cada unidade de ensino que tiver mais de cinqüenta PROFESSORES, a <b>MANTENEDORA</b> assegurará eleição de um Delegado Representante, que terá garantia de emprego e salários a partir da inscrição de sua candidatura até o término do semestre letivo em que sua gestão se encerrar.
<br><b>Parágrafo primeiro</b> – O mandato do Delegado Representanteserá de um ano.
<br><b>Parágrafo segundo</b> – A eleição do Delegado Representanteserá realizada pelo SINPRO na unidade de ensino da MANTENEDORA, por voto direto e secreto. É exigido quorum de 50% (cinqüenta por cento) mais um do corpo docente da unidade onde a eleição ocorrer.
<br><b>Parágrafo terceiro</b> – O SINPRO comunicará a eleição à <b>MANTENEDORA</b> com antecedência mínima de sete dias corridos. Nenhum candidato poderá ser demitido a partir da data da comunicação até o término da apuração.
<br><b>Parágrafo quarto</b> – É condição necessária que os candidatos tenham, à data da eleição, pelo menos um ano de serviço na MANTENEDORA.

<tr><td class=titulo>41. QUADRO DE AVISOS
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> deverá colocar, nas salas de professores, quadro de aviso à disposição do SINPRO para fixação de comunicados de interesse da categoria, sendo vedada a divulgação de matéria político-partidária ou ofensiva a quem quer que seja.

<tr><td class=titulo>42. ASSEMBLÉIAS SINDICAIS
<tr><td class=campo style="text-align:justify">Todo <b>PROFESSOR</b> terá direito a abono de faltas para o comparecimento a assembléias da categoria.
<br><b>Parágrafo primeiro</b> – Na vigência desta Convenção, os abonos estão limitados a dois sábados e mais dois dias úteis para cada período compreendido entre o mês de março e o mês de fevereiro do ano subseqüente. As duas assembléias realizadas durante os dias úteis deverão ocorrer em períodos distintos.
<br><b>Parágrafo segundo</b> – O SINPRO ou a FEPESP deverá informar ao SEMESP ou à MANTENEDORA, por escrito, com antecedência mínima de quinze dias corridos. Na comunicação deverão constar a data e o horário da assembléia.
<br><b>Parágrafo terceiro</b> – Os dirigentes sindicais não estão sujeitos ao limite previsto no parágrafo 1º desta cláusula. As ausências decorrentes do comparecimento às assembléias de suas entidades serão abonadas mediante prévia comunicação formal à MANTENEDORA.
<br><b>Parágrafo quarto</b> – A <b>MANTENEDORA</b> poderá exigir dos PROFESSORES e dos dirigentes sindicais atestado emitido pelo SINPRO ou pela FEPESP que comprove
o seu comparecimento à assembléia.

<tr><td class=titulo>43. CONGRESSOS, SIMPÓSIOS E EQUIVALENTES
<tr><td class=campo style="text-align:justify">Os abonos de falta para comparecimento a congressos e simpósios serão concedidos mediante aceitação por parte da MANTENEDORA, que deverá formalizar por escrito a dispensa do PROFESSOR.
<br><b>Parágrafo único</b> – A participação do <b>PROFESSOR</b> nos eventos descritos no caput não caracterizará atividade extraordinária.

<tr><td class=titulo>44. CONGRESSO DO SINPRO
<tr><td class=campo style="text-align:justify">Em cada ano da vigência desta Convenção, o SINPRO promoverá um evento de natureza política ou pedagógica (congresso ou jornada). A <b>MANTENEDORA</b> abonará as ausências de seus PROFESSORES que participarem do evento, nos seguintes limites:
<blockquote style="margin-top:0;margin-bottom:0">
a) na unidade de ensino que tenha até 49 PROFESSORES será garantido o abono a um PROFESSOR;
<br>b) na unidade de ensino que tenha entre 50 e 99 PROFESSORES será garantido o abono a dois PROFESSORES;
<br>c) na unidade de ensino que tenha mais de cem PROFESSORES será garantido o abono a três PROFESSORES.
</blockquote>
Tais faltas, limitadas ao máximo em dois dias úteis além do sábado, em cada evento, serão abonadas mediante a apresentação de atestado de comparecimento fornecido pelo SINPRO. O <b>PROFESSOR</b> deverá repor as aulas que, por ventura, sejam necessárias para complementação das horas letivas mínimas exigidas pela legislação.

<tr><td class=titulo>45. RELAÇÃO NOMINAL
<tr><td class=campo style="text-align:justify">Na vigência desta Convenção, obriga-se a <b>MANTENEDORA</b> a encaminhar ao SINPRO, até o final do mês de junho de cada ano, a relação nominal dos PROFESSORES que integram seu quadro de funcionários, acompanhada do valor do salário mensal e das guias das contribuições sindical e assistencial.

<tr><td class=titulo>46. FORO CONCILIATÓRIO PARA SOLUÇÃO DE CONFLITOS COLETIVOS
<tr><td class=campo style="text-align:justify">Fica mantida a existência do Foro Conciliatório que tem como objetivo procurar
resolver questões referentes ao não cumprimento de normas estabelecidas na presente Convenção e eventuais divergências trabalhistas existentes entre a <b>MANTENEDORA</b> e seus PROFESSORES.
<br><b>Parágrafo primeiro</b> – O Foro será composto por membros do SEMESP e do SINPRO. As reuniões deverão contar, também, com as partes em conflito que, se assim o desejarem, poderão delegar representantes para substituí-las e/ou serem assistidas por advogados.
<br><b>Parágrafo segundo</b> – O SEMESP e o SINPRO deverão indicar os seus representantes no Foro num prazo de trinta dias a contar da assinatura desta Convenção.
<br><b>Parágrafo terceiro</b> – Cada seção do Foro será realizada no prazo máximo de quinze dias a contar da solicitação formal e obrigatória de qualquer uma das entidades que o compõem, devendo constar na solicitação a data, o local e o horário em que a mesma deverá se realizar. O não-comparecimento de qualquer uma das partes acarretará no encerramento imediato das negociações.
<br><b>Parágrafo quarto</b> – Nenhuma das partes envolvidas ingressará com ação na Justiça do Trabalho durante as negociações de entendimento.
<br><b>Parágrafo quinto</b> – Na ausência de solução do conflito ou na hipótese de não comparecimento de qualquer uma das partes, a comissão responsável pelo Foro fornecerá certidão atestando o encerramento da negociação.
<br><b>Parágrafo sexto</b> – Na hipótese de sucesso das negociações, a critério do Foro, a <b>MANTENEDORA</b> ficará desobrigada de arcar com a multa prevista na cláusula 54 desta Convenção.
<br><b>Parágrafo sétimo</b> – As decisões do Foro terão eficácia legal entre as partes acordantes. O descumprimento das decisões assumidas gerará multa a ser estabelecida no Foro,independentemente daquelas já estabelecidas nesta Convenção.
<br><b>Parágrafo oitavo</b> – Na hipótese de incapacidade econômico-financeira das MANTENEDORAS, os casos serão remetidos para análise e deliberação deste foro.

<tr><td class=titulo>47. COMISSÃO PERMANENTE DE NEGOCIAÇÃO
<tr><td class=campo style="text-align:justify">Fica mantida a Comissão Permanente de Negociação constituída de forma paritária, por três representantes das entidades sindicais profissional e econômica, com o objetivo de:
<blockquote style="margin-top:0;margin-bottom:0">
•fiscalizar o cumprimento das cláusulas vigentes; 
<br>•elucidar eventuais divergências de interpretação das cláusulas desta Convenção; 
<br>•discutir questões não-contempladas na presente Convenção. 
<br>•deliberar no prazo máximo de trinta dias a contar da data da solicitação protocolizada no SEMESP, sobre a isenção prevista na cláusula 36 da presente Convenção; sobre modificação de pagamento da assistência médico-hospitalar, conforme os parágrafos 1º e 3º da cláusula 49 da presente Convenção e sobre o valor da remuneração da hora-aula, conforme o parágrafo 2º da cláusula 14 da presente Convenção. 
<br>•criar subsídios para a Comissão de Tratativas Salariais, através da elaboração de documentos, para a definição das funções/atividades e o regime de trabalho dos PROFESSORES. 
</blockquote>
<b>Parágrafo primeiro</b> - As entidades sindicais componentes da Comissão Permanente deNegociação indicarão seus representantes, no prazo máximo de trinta dias corridos, a contar da assinatura da presente Convenção.
<br><b>Parágrafo segundo</b> - A Comissão Permanente de Negociação deverá reunir-se mensalmente, no décimo dia útil, às 15 horas, alternadamente nas sedes das entidades sindicais que a compõem. No caso específico do item d) do caput, deverá haver convocação específica feita pela entidade sindical patronal.

<tr><td class=titulo>48. ACORDOS INTERNOS – CLÁUSULAS MAIS FAVORÁVEIS
<tr><td class=campo style="text-align:justify">Ficam assegurados os direitos mais favoráveis decorrentes de acordos internos ou de acordos coletivos de trabalho celebrados entre a <b>MANTENEDORA</b> e o SINPRO.

<tr><td class=titulo>49. ASSISTÊNCIA MÉDICO-HOSPITALAR
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> está obrigada a assegurar, a suas expensas, assistência médico-hospitalar a todos os seus PROFESSORES, sendo-lhe facultada a escolha por plano de saúde, seguro-saúde ou convênios com empresas prestadoras de serviços médico-hospitalares. Poderá ainda prestar a referida assistência diretamente, em se tratando de instituições que disponham de serviços de saúde e hospitais próprios ou conveniados. Qualquer que seja a opção, a assistência médico-hospitalar deve assegurar as condições e os requisitos mínimos que seguem relacionados:
<blockquote style="margin-top:0;margin-bottom:0">
1. Abrangência - A assistência médico-hospitalar deve ser realizada no município onde funciona o estabelecimento de ensino superior ou onde vive o PROFESSOR, a critério da MANTENEDORA. Em casos de emergência, deverá haver garantia de atendimento integral em qualquer localidade do Estado de São Paulo ou fixação em contrato, de formas de reembolso.
<br>2. Coberturas mínimas 
<blockquote style="margin-top:0;margin-bottom:0">
2.1 Quarto para quatro pacientes, no máximo.
<br>2.2 Consultas.
<br>2.3 Prazo de internação de 365 dias por ano (comum e UTI/CTI).
<br>2.4 Parto independentemente do estado gravídicio.
<br>2.5 Moléstias infecto-contagiosas que exijam internação.
<br>2.6 Exames laboratoriais, ambulatoriais e hospitalares.
</blockquote>
3. Carência - Não haverá carência na prestação dos serviços médicos laboratoriais.
<br>4. <b>PROFESSOR</b> ingressante - Não haverá carência para o <b>PROFESSOR</b> ingressante, independentemente da data em que for contratado.
<br>5. Pagamento - Caberá ao <b>PROFESSOR</b> o pagamento de 10% (dez por cento) do valor da Assistência Médica, limitado tal pagamento a R$ 8,00, respeitado o disposto nos parágrafos 1º e 2º.
</blockquote>
<b>Parágrafo primeiro</b> – Caso a assistência médico-hospitalar vigente na Instituição venha a sofrer reajuste em virtude de possíveis modificações estabelecidas em legislação que abranja o segmento, Lei 9.656, de 03 de julho de 1998 e MP 2.097-39, de 26 de abril de 2001, ou que vierem a ser estabelecidas em lei, ou por mudança de empresa prestadora de serviço a pedido dos empregados da Instituição, ou por quebra unilateral de contrato por parte da atual empresa prestadora de serviço, a <b>MANTENEDORA</b> continuará a contribuir com o valor mensal vigente até a data da modificação, devendo o <b>PROFESSOR</b> arcar com o valor excedente, que será descontado em folha e consignado no comprovante de pagamento, nos termos do artigo 462 da CLT.
<br><b>Parágrafo segundo</b> - Caso ocorra mudança de empresa prestadora de serviço, por decisão unilateral da MANTENEDORA, com conseqüente reajuste no valor vigente, o <b>PROFESSOR</b> estará isento do pagamento do valor excedente, cabendo à <b>MANTENEDORA</b> prover integralmente a assistência médico-hospitalar, sem nenhum ônus para o PROFESSOR.
<br><b>Parágrafo terceiro</b> - Para efeito do disposto no § 1º desta cláusula, caberá à <b>MANTENEDORA</b> remeter a documentação comprobatória para análise e deliberação da Comissão Permanente de Negociação, nos termos da cláusula 47.
<br><b>Parágrafo quarto</b> - Fica facultado ao <b>PROFESSOR</b> optar pela prestação de assistência médico-hospitalar em uma única instituição de ensino, quando mantiver mais de um vínculo empregatício como PROFESSOR. É necessário que o <b>PROFESSOR</b> se manifeste por escrito, com antecedência mínima de vinte dias, para que a <b>MANTENEDORA</b> possa proceder à suspensão dos serviços.
<br><b>Parágrafo quinto</b> – Caso o <b>PROFESSOR</b> mantenha vínculo empregatício com mais de uma Instituição de Ensino, as MANTENEDORAS, em conjunto, poderão optar por conceder-lhe um único plano de saúde, pago por elas em regime de cotização de custos, respeitadas as condições estabelecidas nesta cláusula.
<br><b>Parágrafo sexto</b> - Mediante pagamento complementar e adesão facultativa devidamente documentada, o <b>PROFESSOR</b> poderá optar pela ampliação dos serviços de saúde garantidos nesta Convenção ou estendê-los a seus dependentes.

<tr><td class=titulo>50. Bolsas de Estudo
<tr><td class=campo style="text-align:justify">Todo <b>PROFESSOR</b> tem direito a bolsas de estudo integrais, incluindo matrícula, para si, seus filhos ou dependentes legais, estes últimos entendidos como aqueles reconhecidos pela legislação do Imposto de Renda ou aqueles que estejam sob a guarda judicial do <b>PROFESSOR</b> e vivam sob sua dependência econômica, devidamente comprovada. Os filhos do <b>PROFESSOR</b> poderão usufruir as bolsas de estudo integrais, sem qualquer ônus, desde que não tenham 25 (vinte e cinco) anos completos ou mais na data de realização do exame vestibular ou do processo seletivo que define o ingresso no curso superior.
<br>As bolsas de estudo são válidas para cursos de graduação, pós-graduação ou seqüenciais existentes e administrados pela <b>MANTENEDORA</b> para a qual o <b>PROFESSOR</b> trabalha, observado o disposto nesta cláusula e parágrafos seguintes.
<br><b>Parágrafo primeiro</b> – O direito às bolsas de estudo passa a vigorar ao término do contrato de experiência, cuja duração não pode exceder de 90 (noventa) dias, conforme parágrafo único do artigo 445 da CLT.
<br><b>Parágrafo segundo</b> - A <b>MANTENEDORA</b> está obrigada a conceder duas bolsas de estudo, sendo que, nos cursos de graduação ou seqüenciais, não será possível que o bolsista conclua mais de um curso nesta condição.
<br><b>Parágrafo terceiro</b> - A utilização do benefício previsto nesta cláusula é transitória e não-habitual e, por isso, não possui caráter remuneratório e nem se vincula, para nenhum efeito, ao salário ou remuneração percebida pelo <b>PROFESSOR</b> nos termos do inciso XIX, do parágrafo 9º do artigo 214 do Decreto 3048, de 06 de maio de 1999 e do parágrafo 2º do artigo 458 da CLT, com a redação dada pela Lei 10243, de 19 de junho de 2001.
<br><b>Parágrafo quarto</b> - As bolsas de estudo serão mantidas quando o <b>PROFESSOR</b> estiver licenciado para tratamento de saúde ou em gozo de licença mediante anuência da MANTENEDORA, excetuado o disposto na cláusula 26 da presente Convenção – Licença sem Remuneração.
<br><b>Parágrafo quinto</b> - No caso de falecimento do PROFESSOR, os dependentes que já se encontram estudando em estabelecimento de ensino superior da <b>MANTENEDORA</b> continuarão a gozar das bolsas de estudo até o final do curso, ressalvado o disposto no parágrafo 8º desta cláusula.
<br><b>Parágrafo sexto</b> - No caso de dispensa sem justa causa durante o período letivo, ficam garantidas ao PROFESSOR, até o final do período letivo, as bolsas de estudo já existentes.
<br><b>Parágrafo sétimo</b> - As bolsas de estudo integrais em cursos de pós-graduação ou especialização existentes e administrados pela <b>MANTENEDORA</b> são válidas exclusivamente para o PROFESSOR, em áreas correlatas às disciplinas que o mesmo ministra na Instituição ou que visem a capacitação docente, respeitados os critérios de seleção exigidos para ingresso no mesmo e obedecerão as seguintes condições:
<blockquote style="margin-top:0;margin-bottom:0">
•os cursos stricto sensu ou de especialização que fixem um número máximo de alunos por turma, são limitadas em 30% (trinta por cento) do total de vagas oferecidas; 
<br>•nos cursos de pós-graduação lato sensu não haverá limites de vagas. Caso a estrutura do curso torne necessária a limitação do número de alunos será observado o disposto na alínea a) deste parágrafo. 
</blockquote>
<b>Parágrafo oitavo</b> - Os bolsistas que forem reprovados no período letivo perderão o direito à bolsa de estudo, voltando a gozar do benefício quando lograrem aprovação no referido período. As disciplinas cursadas em regime de dependência serão de total responsabilidade do bolsista, arcando o mesmo com o seu custo.
<br><b>Parágrafo nono</b> - Considera-se adquirido o direito daquele <b>PROFESSOR</b> que já esteja usufruindo bolsas de estudo em número superior ao definido nesta cláusula.

<tr><td class=titulo>51. AUTORIZAÇÃO PARA DESCONTO EM FOLHA DE PAGAMENTO
<tr><td class=campo style="text-align:justify">O desconto do <b>PROFESSOR</b> em folha de pagamento somente poderá ser realizado mediante sua autorização, nos termos dos artigos 462 e 545 da CLT, quando os valores forem destinados ao custeio de prêmios de seguro, planos de saúde, mensalidades associativas ou outras que constem da sua expressa autorização, desde que não haja previsão expressa de desconto na presente norma coletiva.
<br><b>Parágrafo único</b> – Encontra-se no SINPRO, à disposição da MANTENEDORA, cópia de autorização do <b>PROFESSOR</b> para o desconto da mensalidade associativa.

<tr><td class=titulo>52. ESTABILIDADE PARA PORTADORES DE DOENÇAS GRAVES
<tr><td class=campo style="text-align:justify">Fica assegurada, até alta médica, considerada como apto ao trabalho, ou eventual concessão de aposentadoria por invalidez,estabilidade no emprego aos PROFESSORES acometidos por doenças graves ou incuráveis e aos PROFESSORES portadores do vírus HIV que vierem a apresentar qualquer tipo de infecção ou doença oportunista, resultante da patologia de base.
<br><b>Parágrafo único</b> - São consideradas doenças graves ou incuráveis, a tuberculose ativa, alienação mental, esclerose múltipla, neoplasia maligna, cegueira definitiva, hanseníase, cardiopatia grave, doença de Parkinson, paralisia irreversível e incapacitante, espondiloartrose anquilosante, nefropatia grave, estados do Mal de Paget (osteíte deformante) e contaminação grave por radiação.

<tr><td class=titulo>53. GARANTIAS AO <b>PROFESSOR</b> COM SEQÜELAS E READAPTAÇÃO
<tr><td class=campo style="text-align:justify">Será garantida ao <b>PROFESSOR</b> acidentado no trabalho ou acometido por doença profissional a permanência na empresa em função compatível com o seu estado físico, sem prejuízo na remuneração antes percebida, desde que, após o acidente ou comprovação da aquisição de doença profissional, apresente, cumulativamente, redução da capacidade laboral, atestada pelo órgão oficial e que se tenha tornado incapaz de exercer a função que anteriormente desempenhava, obrigado, porém, o <b>PROFESSOR</b> nessa situação a participar dos processos de readaptação e reabilitação profissional.
<br><b>Parágrafo único</b> – O período de estabilidade do <b>PROFESSOR</b> que se encontre participando dos processos de readaptação e reabilitação profissional será o previsto em lei.

<tr><td class=titulo>54. MULTA POR DESCUMPRIMENTO DA CONVENÇÃO
<tr><td class=campo style="text-align:justify">O descumprimento desta Convenção obrigará a <b>MANTENEDORA</b> ao pagamento de multa correspondente a 1% (um por cento) do salário do PROFESSOR, para cada uma das cláusulas não-cumpridas, acrescidas de juros, a cada <b>PROFESSOR</b> prejudicado.
<br><b>Parágrafo único</b> – A <b>MANTENEDORA</b> está desobrigada de arcar com a multa prevista nesta cláusula, caso o artigo da Convenção já estabeleça uma multa pelo não-cumprimento da mesma.

<tr><td class=titulo>55. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL
<tr><td class=campo style="text-align:justify">Obriga-se a <b>MANTENEDORA</b> a promover o desconto nos exercícios de 2005 e 2006, na folha de pagamento de seus PROFESSORES, sindicalizados e/ou filiados ou não, para recolhimento em favor do SINPRO, entidade legalmente representativa da categoria dos PROFESSORES, na base territorial conferida pela respectiva carta sindical oupelo inciso I, artigo 8º da Constituição Federal, em conta especial, da importância correspondente ao percentual estabelecido ou ao que vier a ser estabelecido na Assembléia Geral da categoria. O recolhimento será realizado obrigatoriamente pela própria MANTENEDORA, em guias próprias, acompanhadas das correspondentes relações nominais e valores devidos. As importâncias destinam-se à criação, manutenção e ampliação dos serviços assistenciais do SINPRO, na conformidade das assembléias gerais.
<br><b>Parágrafo primeiro</b> - Quando a <b>MANTENEDORA</b> deixar de efetuar o recolhimento das contribuições estabelecidas nesta cláusula mediante decisão da referida Assembléia Geral, incorrerá na obrigatoriedade do pagamento de multa, cujo valor corresponderá a 5% (cinco por cento) do total da importância a ser recolhida para o SINPRO, acrescida da parcela correspondente à variação da TR ou de outro índice que vier a substituí-la, a partir do dia seguinte ao vencimento, cabendo à <b>MANTENEDORA</b> a integral responsabilidade pela multa e demais cominações, não podendo as mesmas, de forma alguma, incidir sobre os salários dos PROFESSORES.
<br><b>Parágrafo segundo</b> – Eventuais discordâncias dos PROFESSORES, nos termos do Precedente Normativo nº 74 do TST e da ementa do STF, prolatada nos autos do recurso extraordinário nº 220-700-1, RS, em 06 de outubro de 1998 e publicada no DJ, edição de 13 de novembro de 1998 e do Acórdão de STF, de 07/11/2000, deverão ser comunicadas oficialmente pelo próprio <b>PROFESSOR</b> ao SINPRO, no prazo de 10 dias antes da efetivação do primeiro pagamento, já reajustado, com cópia à MANTENEDORA, sob pena de perderem eficácia.
<br><b>Parágrafo terceiro</b> – O SINPRO encaminhará em tempo hábil ao SEMESP, ata da assembléia geral que fixou a contribuição, os respectivos valores e a época do desconto e do recolhimento.

<tr><td class=titulo>56. NÚCLEO INTERSINDICAL DE CONCILIAÇÃO TRABALHISTA
<tr><td class=campo style="text-align:justify">Fica mantido o Núcleo Intersindical de Conciliação Trabalhista, nos termos previstos pelo artigo 625-C da Consolidação das Leis do Trabalho, com redação dada pela Lei 9.958, de 12 de janeiro de 2000.
<br><b>Parágrafo único</b> – O Núcleo Intersindical de Conciliação Trabalhista terá suas normas definidas pelo SINPRO-SP e pelo SEMESP e fixadas, sob forma de aditamento, à presente Convenção Coletiva.

<tr><td class=campo style="text-align:justify">E por estarem justos e acertados, assinam a presente Convenção Coletiva de Trabalho, a qual será depositada na Delegacia Regional do Trabalho de São Paulo, nos termos do artigo 614 e parágrafos, para fins de arquivo, de modo a surtir, de imediato, os seus efeitos legais.

<tr><td class=campo style="text-align:justify">São Paulo, de junho de 2005. 
<br>
<pre>
<br>Hermes Ferreira Figueiredo              Augusto Cezar Casseb
<br>Presidente do SEMESP                    Presidente do SEMESP São José do Rio Preto
<br>
<br>Celso Napolitano                        Luiz Antonio Barbagli
<br>Presidente da FEPESP                    Presidente do SINPRO – SÃO PAULO
<br>
<br>Solange Cristina Silva Neisy            Martins de Oliveira Cardoso
<br>Presidente do SINPRO – OSASCO           Presidente do SINPRO – JUNDIAÍ
<br>
<br>Idelfonso Paz Dias                      Eduardo de Oliveira
<br>Presidente do SINPRO – SANTOS           Presidente do SINPRO – GUARULHOS 
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