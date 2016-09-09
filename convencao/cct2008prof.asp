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
<title>Convenção Coletiva 2008/9 - Professores</title>
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
<tr><td class=titulo align="center">CONVENÇÃO COLETIVA DE TRABALHO PARA 2008/2009
<tr><td class=titulo align="center">SEMESP
<tr><td class=titulo align="center">PROFESSORES 
<tr><td class=campo style="text-align:justify">Entre as partes, de um lado, o Sindicato dos Professores de São Paulo; Sindicato dos Professores de Campinas e Região (Americana, Amparo, Araras, Campinas, Espírito Santo do Pinhal, Limeira, Mogi-Mirim, Piracicaba e Santa Bárbara D´Oeste); Sindicato dos Professores de Santo André, São Bernardo do Campo e São Caetano do Sul – SINPRO ABC; Sindicato dos Professores de Santos e Região (Bertioga, Cananéia, Caraguatatuba, Cubatão, Eldorado, Guarujá, Iguape, Ilha Bela, Itanhaém, Itariri, Jacupiranga, Juquiá, Miracatu, Mongaguá, Pariquera-Açu, Pedro de Toledo, Peruíbe, Praia Grande, Registro, Santos, São Sebastião, São Vicente, Sete Barras e Ubatuba); Sindicato dos Professores de <b>Osasco</b> e Região (Barueri, Carapicuíba Cotia e Osasco); Sindicato dos Professores de Jundiaí; Sindicato dos Professores de Guarulhos; Sindicato dos Professores de Valinhos e Vinhedo; Sindicato dos Professores de Jaú; Sindicato dos Professores de Indaiatuba, Salto e Itu – SINPRO Vales; Sindicato dos Professores de Sorocaba e Região (Alambari, Alumínio, Angatuba, Apiaí, Araçariguama, Araiçoaba da Serra, Barão de Antonina, Barra do Chapéu, Bofete, Bom Sucesso de Itararé, Buri, Campina do Monte Alegre, Capão Redondo, Cesário Lange, Conchas, Coronel Macedo, Guapiara, Guareí, Ibiúna, Ipero, Itaberá, Itaí, Itaóca, Itapetininga, Itapeva, Itapirapuã Paulista, Itaporanga, Itararé, Mairinque, Nova Campina, Paranapanema, Piedade, Pilar do Sul, Porangaba, Quadra, Riacho Grande, Ribeira, Ribeirão Branco, Ribeirão Grande, Riversul, Salto de Pirapora, São Miguel Arcanjo, São Roque, Sarapuí, Sorocaba, Tapiraí, Taquarituba, Taquarivai, Tatuí, Torre de Pedra, Vargem Grande Paulista, Votorantim) e Sindicato dos Professores de Educação Básica, Superior, Profissionalizantes, livres de Mogi Guaçu e Itapira – SINPRO Guapira; e a Federação dos Professores do Estado de São Paulo – FEPESP, entidades com bases territoriais e representatividades fixadas nas respectivas Cartas Sindicais e no que estabelece o inciso I do artigo 8º da Constituição Federal e de outro, o Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de São Paulo – SEMESP e SEMESP São José do Rio Preto, com representatividade fixada em seus registros sindicais, ao final assinados por seus representantes legais, devidamente autorizados pelas competentes Assembléias Gerais das respectivas categorias, fica estabelecida, nos termos do artigo 611 e seguintes da Consolidação das Leis do Trabalho e do artigo 8º, inciso VI da Constituição Federal, a presente CONVENÇÃO COLETIVA DE TRABALHO:

<tr><td class=titulo>1. Abrangência 
<tr><td class=campo style="text-align:justify">Esta Convenção abrange a categoria econômica dos estabelecimentos particulares de ensino superior no Estado de São Paulo, aqui designados como MANTENEDORA e a categoria profissional diferenciada dos professores, aqui designada simplesmente como PROFESSOR. 
<br><b>Parágrafo único</b> – A categoria dos PROFESSORES abrange todos aqueles que exercem a atividade docente, independentemente da denominação sob a qual a função for exercida. Considera-se atividade docente a função de ministrar aulas. 

<tr><td class=titulo>2. Duração 
<tr><td class=campo style="text-align:justify">Esta Convenção Coletiva de Trabalho terá duração de dois anos, com vigência de 1º de março de 2008 a 28 de fevereiro de 2010. 
<br><b>Parágrafo único</b> – As cláusulas poderão ser reexaminadas na próxima data base, em 1º de março de 2009, em virtude de problemas surgidos na sua aplicação ou do surgimento de normas legais a elas pertinentes, ou em decorrência de aprovação das propostas apresentadas pela Comissão de Aprimoramento das Relações de Trabalho prevista na cláusula 57 da presente Convenção. 

<tr><td class=titulo>3. Reajuste salarial em 2008 
<tr><td class=campo style="text-align:justify">I. Abril de 2008 – A partir de 1º de abril de 2008, será aplicado o reajuste de 4,66% (quatro vírgula sessenta e seis por cento), sobre os salários devidos em 1º de agosto de 2007. Tal reajuste, referente ao mês de abril, deverá ser pago até o 5º dia útil do mês de junho, juntamente com os salários referentes ao mês de maio. 
II. Julho de 2008 – Em 1º de julho de 2008, as MANTENEDORAS deverão aplicar o reajuste de 5,5% (cinco e meio por cento), sobre os salários devidos em 1º de agosto de 2007. 
<br><b>Parágrafo primeiro</b> – Fica estabelecido que o salário de 1º de julho de 2008, reajustado pelo índice definido nesta cláusula, servirá como base de cálculo para a data base de 1º de março de 2009. 

<tr><td class=titulo>4. Reajuste salarial em 1º de março de 2009 
<tr><td class=campo style="text-align:justify">Em 1º de março de 2009, as MANTENEDORAS deverão aplicar sobre os salários devidos em 1º de julho de 2008, o percentual definido pela média aritmética dos índices inflacionários do período compreendido entre 1º de março de 2008 e 28 de fevereiro de 2009, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV), composto com 1,20% (um vírgula vinte por cento). 
<br><b>Parágrafo primeiro</b> – O SEMESP, o SINPRO e a FEPESP comprometem-se a divulgar, em comunicado conjunto, até 20 de março de 2009, o percentual de reajuste salarial calculado pela fórmula definida no caput. 
<br><b>Parágrafo segundo</b> – A base de cálculo para a data-base de 1º de março de 2010 será constituída pelos salários devidos em 1º de julho de 2008, reajustados em 2009 pela fórmula definida no caput. 

<tr><td class=titulo>5. Compensações salariais 
<tr><td class=campo style="text-align:justify">No ano de 2008 será permitida a compensação de eventuais antecipações salariais concedidas no período compreendido entre 1º de março de 2007 e 28 de fevereiro de 2008. Relativamente à data-base de março de 2009 será permitida a compensação de eventuais antecipações salariais concedidas no período compreendido entre 1º de março de 2008 e 28 de fevereiro de 2009.
<br><b>Parágrafo único</b> – Não será permitida, em ambos os casos, a compensação daquelas antecipações salariais que decorrerem de promoções, transferências, ascensão em plano de carreira e os reajustes concedidos com cláusula expressa de não–compensação. 

<tr><td class=titulo>6. Salário do professor ingressante na mantenedora 
<tr><td class=campo style="text-align:justify">A MANTENEDORA não poderá contratar nenhum PROFESSOR por salário inferior ao limite salarial mínimo dos PROFESSORES mais antigos que possuam o mesmo grau de qualificação ou titulação de quem está sendo contratado, respeitado o quadro de carreira da MANTENEDORA. 
<br><b>Parágrafo único</b> – Ao PROFESSOR admitido após 1º de março de 2008 e após 1º de março de 2009, serão concedidos os mesmos percentuais de reajustes e aumentos salariais estabelecidos nas cláusulas 3 e 4, respectivamente, desta norma coletiva. 

<tr><td class=titulo>7. Comprovante de pagamento 
<tr><td class=campo style="text-align:justify">A MANTENEDORA deverá fornecer ao PROFESSOR, mensalmente, comprovante de pagamento, devendo estar discriminados: 
<blockquote style="margin-top:0;margin-bottom:0">a) identificação da MANTENEDORA e do estabelecimento de ensino; 
<br>b) a identificação do PROFESSOR; 
<br>c) a denominação da categoria e, se houver, faixas salariais diferenciadas, inclusive aquelas definidas em eventual plano de carreira da Instituição; 
<br>d) o valor da hora-aula; 
<br>e) a carga horária semanal; 
<br>f) a hora-atividade; 
<br>g) outros eventuais adicionais, inclusive o adicional por tempo de serviço, caso exista; 
<br>h) o descanso semanal remunerado; 
<br>i) as horas extras realizadas; 
<br>j) o valor do recolhimento do FGTS; 
<br>l) o desconto previdenciário; 
<br>m) outros descontos. </blockquote>

<tr><td class=titulo>8. Hora-atividade 
<tr><td class=campo style="text-align:justify">Fica mantido o adicional de 5% (cinco por cento) a título de hora-atividade, destinado exclusivamente ao pagamento do tempo gasto pelo PROFESSOR, fora do estabelecimento de ensino, na preparação de aulas, provas e exercícios, bem como na correção dos mesmos. 

<tr><td class=titulo>9. Adicional noturno 
<tr><td class=campo style="text-align:justify">O trabalho noturno deve ser pago nas atividades realizadas após as 22 horas e corresponde a 25% (vinte e cinco por cento) do valor da hora-aula. 

<tr><td class=titulo>10. Horas extras 
<tr><td class=campo style="text-align:justify">Considera-se atividade extra todo trabalho desenvolvido em horário diferente daquele habitualmente realizado na semana. As atividades extras devem ser pagas com adicional de 100% (cem por cento). 
<br><b>Parágrafo primeiro</b> – Não é considerada atividade extra a participação em cursos de capacitação e aperfeiçoamento docente, desde que aceita livremente pelo PROFESSOR. 
<br><b>Parágrafo segundo</b> – Serão pagas apenas como aulas normais, acrescidas do DSR e da hora-atividade, aquelas que forem adicionadas provisoriamente à carga horária habitual, decorrentes: 
<blockquote style="margin-top:0;margin-bottom:0">a) da substituição temporária de um outro PROFESSOR, com duração predeterminada, decorrente de licença médica, maternidade ou para estudos. Nestes casos, a substituição deverá ser formalizada através de documento firmado entre a MANTENEDORA e o PROFESSOR que aceitar realizá-la; 
<br>b) de substituições eventuais de faltas de PROFESSOR responsável, desde que aceitas livremente pelo PROFESSOR substituto;
<br>c) de reposição de eventuais faltas que foram descontadas dos salários nos meses em que ocorreram; 
<br>d) da realização de cursos eventuais ou de curta duração, inclusive cursos de dependência, e aceitas livremente, mediante documento firmado entre o PROFESSOR convidado a ministrá-los e a MANTENEDORA. 
<br>e) do comparecimento a reuniões didático-pedagógicas, de avaliação e de planejamento, quando realizadas fora de seu horário habitual de trabalho, desde que aceito livremente pelo PROFESSOR. 
</blockquote>
<b>Parágrafo terceiro</b> – A participação em Comissões Internas e Externas da Unidade de Ensino da MANTENEDORA, desde que aceita livremente pelo PROFESSOR mediante documento firmado, será remunerada como aula ou hora normal, acrescida de DSR. 

<tr><td class=titulo>11. Janelas 
<tr><td class=campo style="text-align:justify">Considera-se janela a aula vaga existente no horário do PROFESSOR entre duas outras aulas ministradas no mesmo turno. O pagamento da janela é obrigatório, devendo o PROFESSOR permanecer à disposição da MANTENEDORA neste período, ressalvada a aceitação pelo PROFESSOR, através de acordo formalizado entre as partes antes do início das aulas, quando as janelas não serão pagas. 
<br><b>Parágrafo único</b> - Ocorrendo a hipótese da ressalva supra e caso o PROFESSOR seja solicitado esporadicamente a ministrar aulas ou a desenvolver qualquer outra atividade inerente ao magistério, no horário de janelas não-pagas, essas atividades serão remuneradas como aulas extras, com adicional de 100% (cem por cento). 

<tr><td class=titulo>12. Adicional por atividades em outros municípios 
<tr><td class=campo style="text-align:justify">Quando o PROFESSOR desenvolver suas atividades a serviço da mesma MANTENEDORA em município diferente daquele onde foi contratado e onde ocorre a prestação habitual do trabalho, deverá receber um adicional de 25% (vinte e cinco por cento) sobre o total de sua remuneração no novo município. Quando o PROFESSOR voltar a prestar serviços no município de origem, cessará a obrigação no pagamento do adicional. 
<br><b>Parágrafo primeiro</b> - Nos casos em que ocorrer a transferência definitiva do PROFESSOR, aceita livremente por este, em documento firmado entre as partes, não haverá a incidência do adicional referido no caput, obrigando-se a MANTENEDORA a efetuar o pagamento de um único salário mensal integral, ao PROFESSOR, no ato da transferência, a título de ajuda de custo. 
<br><b>Parágrafo segundo</b> - Fica assegurada a garantia de emprego pelo período de seis meses ao PROFESSOR transferido de município, contados a partir do início do trabalho e/ou da efetivação da transferência. 
<br><b>Parágrafo terceiro</b> – Caso a MANTENEDORA desenvolva atividade acadêmica em municípios considerados conurbados, poderá solicitar isenção do pagamento do adicional determinado no caput, desde que encaminhe material comprobatório ao SEMESP, para análise e deliberação do Foro Conciliatório para Solução de Conflitos Coletivos, previsto na cláusula 47 da presente Convenção.

<tr><td class=titulo>13. Composição do salário mensal do professor 
<tr><td class=campo style="text-align:justify">O salário do PROFESSOR é composto, no mínimo, por três itens: o salário base, o descanso semanal remunerado (DSR) e a hora-atividade. O salário base é calculado pela seguinte equação: número de aulas semanais multiplicado por 4,5 semanas e multiplicado, ainda, pelo valor da hora-aula (artigo 320, parágrafo 1º da CLT). O DSR corresponde a 1/6 (um sexto) do salário base, acrescido, quando houver, do total de horas extras e do adicional noturno (Lei 605/49). A hora-atividade corresponde a 5% (cinco por cento) do total obtido com a somatória de todos os valores acima referidos. 
<br><b>Parágrafo único</b> - A remuneração adicional do PROFESSOR pelo exercício concomitante de função não-docente obedecerá aos critérios estabelecidos entre a MANTENEDORA e o PROFESSOR que aceitar o cargo. 

<tr><td class=titulo>14. Duração da hora-aula 
<tr><td class=campo style="text-align:justify">A duração da hora-aula poderá ser de, no máximo, cinqüenta minutos. 
<br><b>Parágrafo primeiro</b> – Como exceção ao disposto no caput, a hora-aula poderá ter a duração de sessenta minutos nos cursos tecnológicos, desde que tenham sido autorizados ou reconhecidos com essa determinação expressa e cujos PROFESSORES desses cursos tenham sido contratados nessa condição. 
<br><b>Parágrafo segundo</b> – As MANTENEDORAS de Instituições de Ensino que possuam cursos tecnológicos nas condições definidas no parágrafo 1º desta cláusula deverão apresentar à Comissão Permanente de Negociação definida na cláusula 48 da presente Convenção, até o dia 15 de agosto de 2008, a documentação de autorização ou reconhecimento do curso com a determinação expressa de hora-aula com duração de sessenta minutos sob pena de, em não o fazendo, estar sujeita à majoração do valor do salário-aula de acordo com o que estabelece o parágrafo 4º desta cláusula. 
<br><b>Parágrafo terceiro</b> – Caso a Comissão Permanente de Acompanhamento delibere não ter havido determinação expressa do Ministério da Educação para que a duração da hora-aula dos cursos tecnológicos seja de sessenta minutos, a MANTENEDORA deverá majorar o salário-aula de acordo com o que estabelece o parágrafo 4º desta cláusula. 
<br><b>Parágrafo quarto</b> – Em caso de ampliação da duração da hora-aula vigente, respeitado o limite previsto no caput desta cláusula, a MANTENEDORA deverá acrescer ao salário-aula já pago, valor proporcional ao acréscimo do trabalho. 

<tr><td class=titulo>15. Carga horária 
<tr><td class=campo style="text-align:justify">Quando a MANTENEDORA e o PROFESSOR contratarem carga diária de aulas superior aos limites previstos no artigo 318 da CLT, o excedente à carga horária legal será remunerado como aula normal, acrescido de DSR, hora-atividade e vantagens pessoais. 
<br><b>Parágrafo único</b> – Poderá ser flexibilizada a carga horária do PROFESSOR entre jornadas, no exercício de sua função docente e concomitantemente com a atividade administrativa, não havendo assim pagamento, no intervalo, de horas aulas e salários, quando o professor não tenha trabalhado no referido intervalo. 

<tr><td class=titulo>16. Prazo para pagamento de salários
<tr><td class=campo style="text-align:justify">Os salários deverão ser pagos, no máximo, até o quinto dia útil do mês subseqüente ao trabalhado. 
<br><b>Parágrafo único</b> - O não-pagamento dos salários no prazo obriga a MANTENEDORA a pagar multa diária, em favor do PROFESSOR, no valor de 1/50 (um cinqüenta avos) de seu salário mensal. 

<tr><td class=titulo>17. Desconto de faltas 
<tr><td class=campo style="text-align:justify">Na ocorrência de faltas, a MANTENEDORA poderá descontar do salário do PROFESSOR, no máximo, o número de aulas em que o mesmo esteve ausente, o DSR (1/6), a hora-atividade e demais vantagens pessoais proporcionais a estas aulas. 
<br><b>Parágrafo único</b> - É da competência e de integral responsabilidade da MANTENEDORA estabelecer mecanismos de controle de faltas e de pontualidade dos PROFESSORES, conforme a legislação vigente. 

<tr><td class=titulo>18. Atestados médicos e abono de faltas 
<tr><td class=campo style="text-align:justify">A MANTENEDORA está obrigada a abonar as faltas dos PROFESSORES, mediante a apresentação de atestados médicos ou odontológicos. 

<tr><td class=titulo>19. Anotações na carteira de trabalho 
<tr><td class=campo style="text-align:justify">A MANTENEDORA está obrigada a promover, em quarenta e oito horas, as anotações nas Carteiras de Trabalho de seus PROFESSORES, ressalvados eventuais prazos mais amplos permitidos por lei. 
<br><b>Parágrafo único</b> - É obrigatória a anotação na Carteira de Trabalho das mudanças provocadas por ascensão ou alteração de titulação, decorrentes e previstas em plano de carreira. 

<tr><td class=titulo>20. Mudança de disciplina 
<tr><td class=campo style="text-align:justify">O PROFESSOR não poderá ser transferido de uma disciplina para outra, salvo com seu consentimento expresso e por escrito, sob pena de nulidade da referida transferência. 

<tr><td class=titulo>21. Redução de carga horária por extinção ou supressão de disciplina, classe ou turma 
<tr><td class=campo style="text-align:justify">Ocorrendo supressão de disciplina, classe ou turma, em virtude de alteração na estrutura curricular prevista ou autorizada pela legislação vigente ou por dispositivo regimental devidamente aprovado por órgão colegiado da Instituição de Ensino, o PROFESSOR da disciplina classe ou turma deverá ser comunicado da redução da sua carga horária, por escrito, com antecedência mínima de 30 (trinta) dias do início do período letivo e terá prioridade para preenchimento de vaga existente em outra classe ou turma ou em outra disciplina para a qual possua habilitação legal. 
<br><b>Parágrafo primeiro</b> – O PROFESSOR deverá manifestar por escrito, no prazo máximo de 5 (cinco) dias após a comunicação da MANTENEDORA, a não-aceitação da transferência de disciplina ou de classe ou turma ou da redução parcial de sua carga horária. A ausência de manifestação do PROFESSOR caracterizará a sua aceitação. 
<br><b>Parágrafo segundo</b> – Caso o PROFESSOR não aceite a transferência para outra disciplina, classe ou turma ou a redução parcial de carga horária, a MANTENEDORA deverá manter a carga horária semanal existente ou, em caso contrário, proceder à rescisão do contrato de trabalho, por demissão sem justa causa.

<tr><td class=titulo>22. Redução de carga horária por diminuição do número de alunos matriculados 
<tr><td class=campo style="text-align:justify">Na ocorrência de diminuição do número de alunos matriculados que venha a caracterizar a supressão de turmas, curso ou disciplina, o PROFESSOR do curso em questão deverá ser comunicado, por escrito, da redução parcial ou total de sua carga horária até o final da segunda semana de aulas do período letivo. 
<br><b>Parágrafo primeiro</b> - O PROFESSOR deverá manifestar, também por escrito, a aceitação ou não da redução parcial de carga horária no prazo máximo de cinco dias após a comunicação da MANTENEDORA. A ausência de manifestação do PROFESSOR caracterizará a sua não-aceitação. 
<br><b>Parágrafo segundo</b> - Caso o PROFESSOR aceite a redução parcial de carga horária, deverá formalizar documento junto à MANTENEDORA e, em não aceitando, a MANTENEDORA deverá proceder à rescisão do contrato de trabalho, por demissão sem justa causa, caso seja mantida a redução parcial de carga horária. 
<br><b>Parágrafo terceiro</b> - Na hipótese de rescisão contratual, por demissão sem justa causa, o aviso prévio será indenizado, estando a MANTENEDORA desobrigada do pagamento do disposto na cláusula 29 da presente Convenção – Garantia Semestral de Salários. 
<br><b>Parágrafo quarto</b> - Não ocorrendo redução do número de alunos matriculados que venha a caracterizar supressão do curso, de turma ou de disciplina, a MANTENEDORA que reduzir a carga horária do PROFESSOR estará sujeita ao disposto na cláusula 29 desta Convenção – Garantia Semestral de Salários – quando ocorrer a rescisão do contrato de trabalho do PROFESSOR. 

<tr><td class=titulo>23. Abono de faltas por casamento ou luto 
<tr><td class=campo style="text-align:justify">Não serão descontadas, no curso de nove dias corridos, as faltas do PROFESSOR, por motivo de gala ou luto, este em decorrência de falecimento de pai, mãe, filho, cônjuge, companheira (o) e dependente juridicamente reconhecido. 
<br><b>Parágrafo único</b> – Não serão descontadas, no curso de três dias, as faltas do PROFESSOR por motivo de falecimento de sogra, sogro, neto, neta, irmão ou irmão. 

<tr><td class=titulo>24. Irredutibilidade salarial 
<tr><td class=campo style="text-align:justify">É proibida a redução de remuneração mensal ou de carga horária, ressalvada a ocorrência do disposto nas cláusulas 21 e 22 da presente Convenção, ou ainda, quando ocorrer iniciativa expressa do PROFESSOR. Em qualquer hipótese, é obrigatória a concordância recíproca, firmada por escrito. 
<br><b>Parágrafo primeiro</b> – Não havendo concordância recíproca, a parte que deu origem à redução prevista nesta cláusula arcará com a responsabilidade da rescisão contratual. 
<br><b>Parágrafo segundo</b> – Outras atividades, ainda que inerentes ao trabalho docente, que não sejam as de ministrar aulas, de duração temporária e determinada, poderão ser regulamentadas por contrato entre as partes, contendo a caracterização da atividade, o início e a previsão do término. 
<br><b>Parágrafo terceiro</b> – A MANTENEDORA não poderá reduzir o valor da hora-aula dos contratos de trabalho vigentes, ainda que venha a instituir ou modificar plano de carreira. 

<tr><td class=titulo>25. Uniformes 
<tr><td class=campo style="text-align:justify">A MANTENEDORA deverá fornecer gratuitamente dois uniformes por ano, quando o seu uso for exigido.

<tr><td class=titulo>26. Licença sem remuneração 
<tr><td class=campo style="text-align:justify">O PROFESSOR com mais de cinco anos ininterruptos de serviço na MANTENEDORA terá direito a licenciar-se, sem direito à remuneração, por um período máximo de dois anos, não sendo este período de afastamento computado para contagem de tempo de serviço ou para qualquer outro efeito, inclusive legal. 
<br><b>Parágrafo primeiro</b> - A licença ou sua prorrogação deverá ser comunicada por escrito, à MANTENEDORA, com antecedência mínima de noventa dias do período letivo, devendo especificar as datas de início e término do afastamento. A licença só terá início a partir da data expressa no comunicado, mantendo-se, até aí, todas as vantagens contratuais. A intenção de retorno do PROFESSOR à atividade deverá ser comunicada à MANTENEDORA, no mínimo, sessenta dias antes do término do afastamento. 
<br><b>Parágrafo segundo</b> - O término do afastamento deverá coincidir com o início do período letivo. 
<br><b>Parágrafo terceiro</b> - O PROFESSOR que tenha ou exerça cargo de confiança deverá, junto com o comunicado de licença, solicitar seu desligamento do cargo a partir do início do período de licença. 
<br><b>Parágrafo quarto</b> - Considera-se demissionário o PROFESSOR que, ao término do afastamento, não retornar às atividades docentes. 
<br><b>Parágrafo quinto</b> - Ocorrendo a dispensa sem justa causa ao término da licença, o PROFESSOR não terá direito à Garantia Semestral de Salários, prevista na cláusula 29 da presente Convenção. 

<tr><td class=titulo>27. Licença à professora adotante 
<tr><td class=campo style="text-align:justify">Nos termos da Lei 10421, de 15 de abril de 2002, será assegurada licença maternidade à PROFESSORA que vier a adotar ou obtiver guarda judicial de crianças, garantido o emprego no período em que a licença for concedida. 

<tr><td class=titulo>28. Licença paternidade 
<tr><td class=campo style="text-align:justify">A licença paternidade terá duração de cinco dias. 

<tr><td class=titulo>29. Garantia semestral de salários 
<tr><td class=campo style="text-align:justify">Ao PROFESSOR demitido sem justa causa, a MANTENEDORA garantirá: 
<blockquote style="margin-top:0;margin-bottom:0">a) no primeiro semestre, a partir de 1º de janeiro, os salários integrais até o dia 30 de junho; 
<br>b) no segundo semestre, os salários integrais até o dia 31 de dezembro, ressalvado o parágrafo 4º. 
</blockquote>
<b>Parágrafo primeiro</b> - Não terá direito à Garantia Semestral de Salários o PROFESSOR que, na data da comunicação da dispensa, contar com menos de 18 (dezoito) meses de serviço prestado à MANTENEDORA, ressalvado o parágrafo 4º desta cláusula. 
<br><b>Parágrafo segundo</b> – No caso de demissões efetuadas no final do primeiro semestre letivo, para não ficar obrigada a pagar ao PROFESSOR os salários do segundo semestre, a MANTENEDORA deverá observar as seguintes disposições: 
<blockquote style="margin-top:0;margin-bottom:0">a) com aviso prévio a ser trabalhado, a demissão deverá ser formalizada com antecedência mínima de trinta dias do início das férias; 
<br>b) sendo o aviso prévio indenizado, a demissão deverá ser formalizada até um dia antes do início das férias, ainda que as férias tenham seu início programado para o mês de julho, obedecendo ao que dispõe a cláusula 39 da presente Convenção.
</blockquote>
<b>Parágrafo terceiro</b> - No caso de demissões efetuadas no final do ano letivo, para não ficar obrigada a pagar ao PROFESSOR os salários do primeiro semestre do ano seguinte, a MANTENEDORA deverá observar as seguintes disposições: 
<blockquote style="margin-top:0;margin-bottom:0">a) com aviso prévio a ser trabalhado, a demissão deverá ser formalizada com antecedência mínima de trinta dias do início do recesso escolar; 
<br>b) sendo o aviso prévio indenizado, a demissão deverá ser formalizada até um dia antes do início do recesso escolar. 
<br>Os dias de aviso prévio que forem indenizados não contarão como tempo de serviço para efeito do pagamento da Garantia Semestral de Salários, conforme o estabelecido nesta cláusula. 
</blockquote>
<b>Parágrafo quarto</b> - Quando as demissões ocorrerem a partir de 16 de outubro, a MANTENEDORA pagará, independentemente do tempo de serviço do PROFESSOR, valor correspondente à remuneração devida até o dia 18 de janeiro do ano subseqüente, inclusive, ressalvados os contratos de experiência e por prazo determinado, estes últimos válidos somente nos casos de substituição temporária, conforme o disposto na alínea a) do parágrafo 2º da cláusula 10 da presente Convenção, não sendo devido o pagamento acumulativo de aviso-prévio. 
<br><b>Parágrafo quinto</b> – Na vigência da presente Convenção os PROFESSORES serão remunerados a partir da data de início de suas atividades na MANTENEDORA, incluindo o período de planejamento escolar. 
<br><b>Parágrafo sexto</b> - Os salários complementares previstos nesta cláusula terão natureza indenizatória, não integrando, para nenhum efeito legal, o tempo de serviço do PROFESSOR. 
<br><b>Parágrafo sétimo</b> - O aviso prévio de trinta dias previsto no artigo 487 da CLT já está integrado às indenizações tratadas nesta cláusula. 

<tr><td class=titulo>30. Pedido de demissão em final de ano letivo 
<tr><td class=campo style="text-align:justify">O PROFESSOR que, no final do ano letivo, comunicar sua demissão até o dia que antecede o início do recesso escolar, será dispensado do cumprimento do aviso prévio e terá direito a receber, como indenização, a remuneração até o dia 18 de janeiro do ano subseqüente, independentemente do tempo de serviço na MANTENEDORA. 

<tr><td class=titulo>31. Garantia de emprego à gestante 
<tr><td class=campo style="text-align:justify">É proibida a dispensa arbitrária ou sem justa causa da PROFESSORA gestante, desde o início da gravidez até sessenta dias após o término do afastamento legal. O aviso prévio começará a contar a partir do término do período de estabilidade. 

<tr><td class=titulo>32. Creches 
<tr><td class=campo style="text-align:justify">É obrigatória a instalação de local destinado à guarda de crianças de até seis meses, quando a MANTENEDORA mantiver contratada, em jornada integral, pelo menos trinta funcionárias com idade superior a 16 anos. A manutenção da creche poderá ser substituída pelo pagamento do reembolso-creche, nos termos da legislação em vigor (artigo 389, parágrafo 1º da CLT e Portarias MTb nº 3296 de 03.09.86 e nº670, de 27/08/97), ou ainda, a celebração de convênio com uma entidade reconhecidamente idônea.

<tr><td class=titulo>33. Garantias ao professor em vias de aposentadoria 
<tr><td class=campo style="text-align:justify">Fica assegurado ao PROFESSOR que, comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria integral por tempo de serviço ou da aposentadoria por idade, a garantia de emprego durante o período que faltar até a aquisição do direito. 
<br><b>Parágrafo primeiro</b> – A garantia de emprego é devida ao PROFESSOR que estiver contratado pela MANTENEDORA há pelo menos três anos. 
<br><b>Parágrafo segundo</b> – A comprovação à MANTENEDORA deverá ser feita mediante a apresentação de documento que ateste o tempo de serviço. Este documento deverá ser emitido por pessoa credenciada junto ao órgão previdenciário. Se o PROFESSOR depender de documentação para realização da contagem, terá um prazo de trinta dias, a contar da data prevista ou marcada para homologação da rescisão contratual. Comprovada a solicitação de tal documentação, os prazos serão prorrogados até que a mesma seja emitida, assegurando-se, nessa situação, o pagamento dos salários pelo prazo máximo de cento e vinte dias. 
<br><b>Parágrafo terceiro</b> – O contrato de trabalho do PROFESSOR só poderá ser rescindido por mútuo acordo homologado pelo SINPRO ou pedido de demissão. 
<br><b>Parágrafo quarto</b> – Havendo acordo formal entre as partes, o PROFESSOR poderá exercer outra função, inerente ao magistério, durante o período em que estiver garantido pela estabilidade. 
<br><b>Parágrafo quinto</b> – O aviso prévio, em caso de demissão sem justa causa, integra o período de estabilidade previsto nesta cláusula. 
<br><b>Parágrafo sexto</b> – Para garantir a estabilidade prevista nesta cláusula, o PROFESSOR deverá encaminhar à MANTENEDORA, dentro da prorrogação prevista no parágrafo 2º, documentação que demonstre a tramitação do processo que atesta o tempo de serviço. 

<tr><td class=titulo>34. Multa por atraso na homologação 
<tr><td class=campo style="text-align:justify">A MANTENEDORA deve pagar as verbas devidas na rescisão contratual no dia seguinte ao término do aviso prévio, quando trabalhado, ou dez dias após o desligamento, quando houver dispensa do cumprimento de aviso prévio. O atraso no pagamento das verbas rescisórias obrigará a MANTENEDORA ao pagamento de multa, em favor do PROFESSOR, correspondente a um mês de sua remuneração, conforme o disposto no parágrafo 8º do artigo 477 da CLT. A partir do vigésimo dia de atraso da homologação da rescisão, a contar da data estabelecida pela legislação para o pagamento das verbas rescisórias, a MATENEDORA estará obrigada, ainda, a pagar ao PROFESSOR multa diária de 0,2% (dois décimos percentuais) do salário mensal. A MANTENEDORA estará desobrigada de pagar a referida multa quando o atraso da homologação vier a ocorrer, comprovadamente, por motivos alheios a sua vontade. 
<br><b>Parágrafo único</b> – O SINPRO está obrigado a fornecer comprovante de comparecimento sempre que a MANTENEDORA se apresentar para homologação das rescisões contratuais e comprovar a convocação do PROFESSOR. 

<tr><td class=titulo>35. Demissão por justa causa 
<tr><td class=campo style="text-align:justify">Quando houver demissão por justa causa, nos termos do art. 482 da CLT, a MANTENEDORA está obrigada a determinar na carta-aviso o motivo que deu origem à dispensa. Caso contrário, fica descaracterizada a justa causa. 

<tr><td class=titulo>36. Readmissão do professor
<tr><td class=campo style="text-align:justify">O PROFESSOR que for readmitido até doze meses após o seu desligamento ficará desobrigado de firmar contrato de experiência. 

<tr><td class=titulo>37. Indenizações por dispensa imotivada 
<tr><td class=campo style="text-align:justify">O PROFESSOR demitido sem justa causa terá direito a uma indenização, além do aviso prévio legal de trinta dias e das indenizações previstas na cláusula 29 desta Convenção, quando forem devidas, nas condições abaixo especificadas: 
<blockquote style="margin-top:0;margin-bottom:0">a) três (03) dias para cada ano trabalhado na MANTENEDORA; 
<br>b) aviso prévio adicional de quinze dias, caso o PROFESSOR tenha, no mínimo, cinqüenta anos de idade e que, à data do desligamento, conte com pelo menos um ano de serviço na MANTENEDORA. 
</blockquote>
<b>Parágrafo primeiro</b> – Não terá direito à indenização assegurada na alínea a) do caput o PROFESSOR que tiver recebido, durante pelo menos um ano, pagamento mensal de adicional por tempo de serviço decorrente de plano de cargos e salários ou de anuênio, qüinqüênio ou equivalente, cujo valor corresponda a, no mínimo, 1% (um por cento) do valor da hora-aula por ano trabalhado e, por conseqüência, do salário mensal. A MANTENEDORA deverá apresentar, no momento da homologação, documentos que comprovem o pagamento ao PROFESSOR do referido adicional por tempo de serviço. 
<br><b>Parágrafo segundo</b> – Não terá direito à indenização assegurada na alínea b) do caput, o PROFESSOR que, na data de admissão na MANTENEDORA, contar com mais de cinqüenta anos de idade. 
<br><b>Parágrafo terceiro</b> – O pagamento das verbas indenizatórias previstas nesta cláusula não será cumulativo, cabendo ao PROFESSOR, no desligamento, o maior valor monetário entre os previstos nas alíneas a) e b) do caput. 
<br><b>Parágrafo quarto</b> – Essas indenizações não contarão, para nenhum efeito, como tempo de serviço. 

<tr><td class=titulo>38. Atestados de afastamento e salários 
<tr><td class=campo style="text-align:justify">Sempre que solicitada, a MANTENEDORA deverá fornecer ao PROFESSOR atestado de afastamento e salário (AAS), previsto na legislação previdenciária. 

<tr><td class=titulo>39. Férias 
<tr><td class=campo style="text-align:justify">As férias anuais dos PROFESSORES serão coletivas, com duração de trinta dias corridos e gozados em julho de 2008 e julho de 2009. Qualquer alteração deverá ser aprovada por órgão competente, conforme o estabelecido em Estatuto ou Regimento e deverá constar do calendário escolar. 
<br><b>Parágrafo primeiro</b> – A MANTENEDORA está obrigada a pagar o salário das férias e o abono constitucional de 1/3 (um terço) até quarenta e oito horas antes do início das férias. 
<br><b>Parágrafo segundo</b> – As férias não poderão ser iniciadas aos domingos, feriados, dias de compensação do descanso semanal remunerado e nem aos sábados, quando estes não forem dias normais de aula. 
<br><b>Parágrafo terceiro</b> – Também terá direito às férias coletivas de trinta dias corridos nos períodos estabelecidos no caput, O PROFESSOR que, além de ministrar aulas, tenha cargo de confiança ou exerça outras atividades na MANTENEDORA.
Caso o exercício da atividade administrativa impossibilite a concessão de férias nos termos do caput, as férias anuais desse PROFESSOR poderão ser gozadas em dois períodos, um deles obrigatoriamente no mês de julho de cada ano. 
<br><b>Parágrafo quarto</b> – Na hipótese da divisão das férias anuais do PROFESSOR nos termos do parágrafo anterior, um dos períodos não poderá ser inferior a 10 (dez) dias, sendo proibido o exercício de qualquer atividade nesses períodos. 

<tr><td class=titulo>40. Recesso escolar 
<tr><td class=campo style="text-align:justify">O recesso escolar anual é obrigatório e tem duração de trinta dias corridos, gozados preferencialmente no mês de janeiro de cada ano. Durante o recesso escolar anual que não pode, de maneira alguma, coincidir com o período definido para as férias coletivas do ano respectivo, o PROFESSOR não poderá ser convocado para nenhum trabalho. 
<br><b>Parágrafo primeiro</b> – Na vigência da presente Convenção, as instituições cujos calendários escolares, determinados pelo órgão competente conforme o estabelecido em Estatuto ou Regimento, não observarem o determinado pelo caput para o recesso escolar anual dos PROFESSORES, poderão concedê-lo em um período de, no mínimo vinte dias corridos, e em, no máximo, mais dois períodos com igual número de dias corridos, desde que observem as seguintes condições: 
<blockquote style="margin-top:0;margin-bottom:0">a) Vinte dias corridos em janeiro de 2008 e os dois períodos com igual número de dias corridos, obrigatoriamente no período compreendido entre março de 2008 e fevereiro de 2009. 
<br>b) Vinte dias corridos em janeiro de 2009 e os dois períodos com igual número de dias corridos, obrigatoriamente no período compreendido entre março de 2009 e fevereiro de 2010. 
</blockquote>
<b>Parágrafo segundo</b> – No caso dos calendários escolares preverem a divisão do recesso escolar dos PROFESSORES, os períodos definidos na conformidade do parágrafo primeiro não poderão ser iniciados aos domingos, feriados, dias de compensação do descanso semanal remunerado e nem aos sábados, quando estes não forem dias normais de aulas. 
<br><b>Parágrafo terceiro</b> – As Instituições cujas atividades não possam ser interrompidas, tais como aquelas desenvolvidas em hospital, clínica, laboratório de análise, escritórios experimentais, pesquisas, dentre outros, ou que ministrem cursos em que sejam utilizadas instalações específicas ou que prestem atendimento à comunidade que não pode ser suspenso, poderão conceder aos PROFESSORES o recesso escolar anual definido no caput de maneira escalonada ao longo de cada ano. 
<br><b>Parágrafo quarto</b> – Os calendários escolares que definirão os períodos de recesso escolar dos PROFESSORES serão obrigatoriamente divulgados aos PROFESSORES até o início de cada período letivo e enviados ao SINPRO. 

<tr><td class=titulo>41. Delegado representante 
<tr><td class=campo style="text-align:justify">A MANTENEDORA que tiver mais de 50 (cinqüenta) PROFESSORES assegurará eleição de Delegados Representantes, com mandato de 1 (um) ano, que terão garantia de emprego e salários a partir da inscrição de sua candidatura até o término do semestre letivo em que sua gestão se encerrar, nos seguintes limites: 
<blockquote style="margin-top:0;margin-bottom:0">a) Na MANTENEDORA que tenha até 100 (cem) PROFESSORES, será garantida a eleição de 1 (um) delegado representante; 
<br>b) Na MANTENEDORA que tenha mais de 200 (duzentos) PROFESSORES, será garantida a eleição de 2 (dois) delegados representantes; 
</blockquote>
<b>Parágrafo primeiro</b> – O mandato dos Delegados Representantes será de um ano. 
<br><b>Parágrafo segundo</b> – A eleição dos Delegados Representantes será realizada pelo SINPRO nas unidades de ensino da MANTENEDORA, por voto direto e secreto. É exigido quorum de 50% (cinqüenta por cento) mais um do corpo docente da unidade onde a eleição ocorrer. 
<br><b>Parágrafo terceiro</b> – O SINPRO comunicará a eleição à MANTENEDORA, com a relação dos candidatos inscritos, com antecedência mínima de sete dias corridos, da data da eleição. Nenhum candidato poderá ser demitido a partir da data da comunicação até o término da apuração. 
<br><b>Parágrafo quarto</b> – É condição necessária que os candidatos sejam filiados ao SINPRO e que tenham, à data da eleição, pelo menos um ano de serviço na MANTENEDORA. 

<tr><td class=titulo>42. Quadro de avisos 
<tr><td class=campo style="text-align:justify">A MANTENEDORA deverá colocar, nas salas de professores, quadro de aviso à disposição do SINPRO para fixação de comunicados de interesse da categoria, sendo vedada a divulgação de matéria político-partidária ou ofensiva a quem quer que seja. 
<br><b>Parágrafo único</b> – O dirigente sindical terá livre acesso à sala dos PROFESSORES, no horário de intervalo das aulas, para atualização do material divulgado no quadro de avisos, uma única vez em cada mês. 

<tr><td class=titulo>43. Assembléias sindicais 
<tr><td class=campo style="text-align:justify">Todo PROFESSOR terá direito a abono de faltas para o comparecimento a assembléias da categoria. 
<br><b>Parágrafo primeiro</b> - Na vigência desta Convenção, os abonos estão limitados a dois sábados e mais dois dias úteis para cada período compreendido entre o mês de março e o mês de fevereiro do ano subseqüente. As duas assembléias realizadas durante os dias úteis deverão ocorrer em períodos distintos. 
<br><b>Parágrafo segundo</b> - O SINPRO ou a FEPESP deverá informar ao SEMESP ou à MANTENEDORA, por escrito, com antecedência mínima de quinze dias corridos. Na comunicação deverão constar a data e o horário da assembléia. 
<br><b>Parágrafo terceiro</b> - Os dirigentes sindicais não estão sujeitos ao limite previsto no parágrafo 1º desta cláusula. As ausências decorrentes do comparecimento às assembléias de suas entidades serão abonadas mediante prévia comunicação formal à MANTENEDORA. 
<br><b>Parágrafo quarto</b> - A MANTENEDORA poderá exigir dos PROFESSORES e do dirigente sindical atestado emitido pelo SINPRO ou pela FEPESP que comprove o seu comparecimento à assembléia. 

<tr><td class=titulo>44. Congressos, simpósios e equivalentes 
<tr><td class=campo style="text-align:justify">Os abonos de falta para comparecimento a congressos e simpósios serão concedidos mediante aceitação por parte da MANTENEDORA, que deverá formalizar por escrito a dispensa do PROFESSOR. 
<br><b>Parágrafo único</b> - A participação do PROFESSOR nos eventos descritos no caput não caracterizará atividade extraordinária. 

<tr><td class=titulo>45. Congresso do Sinpro
<tr><td class=campo style="text-align:justify">Em cada ano de vigência desta Convenção, o SINPRO promoverá um evento de natureza política ou pedagógica (congresso ou jornada). A MANTENEDORA abonará as ausências de seus PROFESSORES que participarem do evento, nos seguintes limites: 
<blockquote style="margin-top:0;margin-bottom:0">a) na unidade de ensino que tenha até 49 (quarenta e nove) PROFESSORES será garantido o abono a 1 (um) PROFESSOR; 
<br>b) na unidade de ensino que tenha entre 50 (cinqüenta) e 99 (noventa e nove) PROFESSORES será garantido o abono a 2 (dois) PROFESSORES; 
<br>c) na unidade de ensino que tenha mais de 100 (cem) PROFESSORES será garantido o abono a 3 (três) PROFESSORES. 
</blockquote>
Tais faltas, limitadas ao máximo em dois dias úteis além do sábado, em cada evento, serão abonadas mediante a apresentação de atestado de comparecimento fornecido pelo SINPRO. O PROFESSOR deverá repor as aulas que, por ventura, sejam necessárias para complementação das horas letivas mínimas exigidas pela legislação. 

<tr><td class=titulo>46. Relação nominal 
<tr><td class=campo style="text-align:justify">Na vigência desta Convenção, obriga-se a MANTENEDORA a encaminhar ao SINPRO, até o final do mês de junho de cada ano, a relação nominal dos PROFESSORES que integram seu quadro de funcionários, acompanhada do valor do salário mensal e das guias das contribuições sindical e assistencial. 

<tr><td class=titulo>47. Foro Conciliatório para Solução de Conflitos Coletivos 
<tr><td class=campo style="text-align:justify">Fica mantida a existência do Foro Conciliatório que tem como objetivo procurar resolver questões referentes ao não cumprimento de normas estabelecidas na presente Convenção e eventuais divergências trabalhistas existentes entre a MANTENEDORA e seus PROFESSORES. 
<br><b>Parágrafo primeiro</b> - O Foro será composto por membros do SEMESP e do SINPRO. As reuniões deverão contar, também, com as partes em conflito que, se assim o desejarem, poderão delegar representantes para substituí-las e/ou serem assistidas por advogados. 
<br><b>Parágrafo segundo</b> - O SEMESP e o SINPRO deverão indicar os seus representantes no Foro num prazo de trinta dias a contar da assinatura desta Convenção. 
<br><b>Parágrafo terceiro</b> - Cada seção do Foro será realizada no prazo máximo de quinze dias a contar da solicitação formal e obrigatória de qualquer uma das entidades que o compõem, devendo constar na solicitação a data, o local e o horário em que a mesma deverá se realizar. O não-comparecimento de qualquer uma das partes acarretará no encerramento imediato das negociações. 
<br><b>Parágrafo quarto</b> - Nenhuma das partes envolvidas ingressará com ação na Justiça do Trabalho durante as negociações de entendimento. 
<br><b>Parágrafo quinto</b> - Na ausência de solução do conflito ou na hipótese de não-comparecimento de qualquer uma das partes, a comissão responsável pelo Foro fornecerá certidão atestando o encerramento da negociação. 
<br><b>Parágrafo sexto</b> - Na hipótese de sucesso das negociações, a critério do Foro, a MANTENEDORA ficará desobrigada de arcar com a multa prevista na cláusula 55 desta Convenção. 
<br><b>Parágrafo sétimo</b> - As decisões do Foro terão eficácia legal entre as partes acordantes. O descumprimento das decisões assumidas gerará multa a ser estabelecida no Foro, independentemente daquelas já estabelecidas nesta Convenção.
<br><b>Parágrafo oitavo</b> – Na hipótese de incapacidade econômico-financeira das MANTENEDORAS, os casos serão remetidos para análise e deliberação deste foro. 

<tr><td class=titulo>48. Comissão Permanente de Negociação 
<tr><td class=campo style="text-align:justify">Fica mantida a Comissão Permanente de Negociação constituída de forma paritária, por três representantes das entidades sindicais (profissional e econômica), com o objetivo de: 
<blockquote style="margin-top:0;margin-bottom:0">a) fiscalizar o cumprimento das cláusulas vigentes; 
<br>b) elucidar eventuais divergências de interpretação das cláusulas desta Convenção; 
<br>c) discutir questões não-contempladas na presente Convenção. 
<br>d) deliberar no prazo máximo de trinta dias a contar da data da solicitação protocolizada no SEMESP, sobre modificação de pagamento da assistência médico-hospitalar, conforme os parágrafos 1º e 3º da cláusula 50 da presente Convenção e sobre o valor da remuneração da hora-aula, conforme o parágrafo 2º da cláusula 14 da presente Convenção. 
<br>e) criar subsídios para a Comissão de Tratativas Salariais, através da elaboração de documentos, para a definição das funções/atividades e o regime de trabalho dos PROFESSORES. 
</blockquote>
<b>Parágrafo primeiro</b> - As entidades sindicais componentes da Comissão Permanente de Negociação indicarão seus representantes, no prazo máximo de trinta dias corridos, a contar da assinatura da presente Convenção. 
<br><b>Parágrafo segundo</b> - A Comissão Permanente de Negociação deverá reunir-se mensalmente, no décimo dia útil, às 15 horas, alternadamente nas sedes das entidades sindicais que a compõem. No caso específico do item “d“ do caput, deverá haver convocação específica feita pela entidade sindical patronal. 

<tr><td class=titulo>49. Acordos internos - cláusulas mais favoráveis 
<tr><td class=campo style="text-align:justify">Ficam assegurados os direitos mais favoráveis decorrentes de acordos internos ou de acordos coletivos de trabalho celebrados entre a MANTENEDORA e o SINPRO. 

<tr><td class=titulo>50. Assistência médico-hospitalar 
<tr><td class=campo style="text-align:justify">A MANTENEDORA está obrigada a assegurar, às suas expensas, nos limites estabelecidos nesta cláusula, assistência médico-hospitalar a todos os seus PROFESSORES, sendo-lhe facultada a escolha por plano de saúde, seguro-saúde ou convênios com empresas prestadoras de serviços médico-hospitalares. Poderá ainda prestar a referida assistência diretamente, em se tratando de instituições que disponham de serviços de saúde e hospitais próprios ou conveniados. Qualquer que seja a opção feita, a assistência médico-hospitalar deve assegurar as condições e os requisitos mínimos que seguem relacionados: 
<br><b>1.Abrangência </b>
<blockquote style="margin-top:0;margin-bottom:0">A assistência médico-hospitalar deve ser realizada no município onde funciona o estabelecimento de ensino superior ou onde vive o PROFESSOR, a critério da MANTENEDORA. Em casos de emergência, deverá haver garantia de atendimento integral em qualquer localidade do Estado de São Paulo ou fixação, em contrato, de formas de reembolso. 
</blockquote>
<b>2. Coberturas mínimas </b>
<blockquote style="margin-top:0;margin-bottom:0">2.1 Quarto para quatro pacientes, no máximo. 
<br>2.2 Consultas.
<br>2.3 Prazo de internação de 365 dias por ano (comum e UTI/CTI) 
<br>2.4 Parto, independentemente do estado gravídico. 
<br>2.5 Moléstias infecto-contagiosas que exijam internação. 
<br>2.6 Exames laboratoriais, ambulatoriais e hospitalares. 
</blockquote>
<b>3. Carência </b>
<blockquote style="margin-top:0;margin-bottom:0">Não haverá carência na prestação dos serviços médicos e laboratoriais. 
</blockquote>
<b>4. Professor ingressante </b>
<blockquote style="margin-top:0;margin-bottom:0">Não haverá carência para o PROFESSOR ingressante, independentemente do mês em que for contratado. 
</blockquote>
<b>5. Pagamento </b>
<blockquote style="margin-top:0;margin-bottom:0">Caberá ao PROFESSOR o pagamento de 10% (dez por cento) do valor da Assistência Médica, respeitado o disposto nos parágrafos 1º, 2º e 3º. 
</blockquote>
<b>Parágrafo primeiro</b> – A MANTENEDORA deverá enviar ao SINPRO cópia do contrato formalizado com a empresa de assistência médico–hospitalar ou de seguro saúde ou de medicina de grupo que comprove o valor pago. 
<br><b>Parágrafo segundo</b> – Caso a assistência médico-hospitalar vigente na Instituição venha a sofrer reajuste em virtude de possíveis modificações estabelecidas em legislação que abranja o segmento - Lei 9.656, de 03 de junho de 1998 e MP 2.097-39, de 26 de abril de 2001, ou que vierem a ser estabelecidas em lei, ou por mudança de empresa prestadora de serviço, a pedido dos empregados da Instituição ou por quebra de contrato, unilateralmente, por parte da atual empresa prestadora de serviço, a MANTENEDORA continuará a contribuir com o valor mensal vigente até a data da modificação, devendo o PROFESSOR arcar com o valor excedente, que será descontado em folha e consignado no comprovante de pagamento, nos termos do artigo 462 da CLT. 
<br><b>Parágrafo terceiro</b> – Caso ocorra mudança de empresa prestadora de serviço, por decisão unilateral da MANTENEDORA, com conseqüente reajuste no valor vigente, o PROFESSOR estará isento do pagamento do valor excedente, cabendo à MANTENEDORA prover integralmente a assistência médico-hospitalar, sem nenhum ônus para o PROFESSOR. 
<br><b>Parágrafo quarto</b> – Para efeito do disposto no parágrafo primeiro desta cláusula, caberá à MANTENEDORA remeter a documentação comprobatória para análise e deliberação da Comissão Permanente de Negociação. 
<br><b>Parágrafo quinto</b> – Fica facultado ao PROFESSOR optar pela prestação de assistência médico-hospitalar em uma única instituição de ensino, quando mantiver mais de um vínculo empregatício como PROFESSOR. É necessário que o PROFESSOR se manifeste por escrito, com antecedência mínima de vinte dias, para que a MANTENEDORA possa proceder à suspensão dos serviços.
<br><b>Parágrafo sexto</b> – Caso o PROFESSOR mantenha vínculo empregatício com mais de uma Instituição de Ensino, as MANTENEDORAS, em conjunto, poderão optar por conceder-lhe um único plano de saúde, pago por elas, em regime de cotização de custos, respeitadas as condições estabelecidas nesta cláusula. 
<br><b>Parágrafo sétimo</b> – Mediante pagamento complementar e adesão facultativa, devidamente documentada, o PROFESSOR poderá optar pela ampliação dos serviços de saúde garantidos nesta Convenção ou estendê-los a seus dependentes. 

<tr><td class=titulo>51. Bolsas de estudo 
<tr><td class=campo style="text-align:justify">Todo PROFESSOR tem direito a bolsas de estudo integrais, incluindo matrícula, para si, seus filhos ou dependentes legais, estes últimos entendidos como aqueles reconhecidos pela legislação do Imposto de Renda ou aqueles que estejam sob a guarda judicial do PROFESSOR e vivam sob sua dependência econômica, devidamente comprovada. Os filhos do PROFESSOR poderão usufruir as bolsas de estudo integrais, sem qualquer ônus, desde que não tenham vinte e cinco anos completos ou mais na data da efetivação da matrícula no curso superior. As bolsas de estudo são válidas para cursos de graduação, pós-graduação ou seqüenciais existentes e administrados pela MANTENEDORA para a qual o PROFESSOR trabalha, observado o disposto nesta cláusula e parágrafos seguintes. 
<br><b>Parágrafo primeiro</b> – O direito às bolsas de estudo passa a vigorar ao término do contrato de experiência, cuja duração não pode exceder de 90 (noventa) dias, conforme parágrafo único do artigo 445 da CLT. 
<br><b>Parágrafo segundo</b> - A MANTENEDORA está obrigada a conceder, no máximo, duas bolsas de estudo, sendo que, nos cursos de graduação ou seqüenciais, não será possível que o bolsista conclua mais de um curso nessa condição. 
<br><b>Parágrafo terceiro</b> – A utilização do benefício previsto nesta cláusula, caracterizada como doação por não impor qualquer contraprestação de serviços, é transitória e não habitual e, por isso, não possui caráter remuneratório e nem se vincula, para nenhum efeito, ao salário ou remuneração percebida pelo PROFESSOR, nos termos do inciso XIX, do parágrafo 9º do artigo 214 do Decreto 3.048, de 06 de maio de 1999 e da Lei 10.243, de 19 de junho de 2001 e visa a capacitação dos beneficiários. 
<br><b>Parágrafo quarto</b> - As bolsas de estudo serão mantidas quando o PROFESSOR estiver licenciado para tratamento de saúde ou em gozo de licença mediante anuência da MANTENEDORA, excetuado o disposto na cláusula 26 da presente Convenção – Licença sem Remuneração. 
<br><b>Parágrafo quinto</b> - No caso de falecimento do PROFESSOR, os dependentes que já se encontram estudando em estabelecimento de ensino superior da MANTENEDORA continuarão a gozar das bolsas de estudo até o final do curso, ressalvado o disposto no parágrafo 8º desta cláusula. 
<br><b>Parágrafo sexto</b> - No caso de dispensa sem justa causa durante o período letivo, ficam garantidas ao PROFESSOR, até o final do período letivo, as bolsas de estudo já existentes. 
<br><b>Parágrafo sétimo</b> - As bolsas de estudo integrais em cursos de pós-graduação ou especialização existentes e administrados pela MANTENEDORA são válidas exclusivamente para o PROFESSOR, em áreas correlatas às disciplinas que o mesmo ministra na Instituição e que visem a capacitação docente, respeitados os critérios de seleção exigidos para ingresso no mesmo e obedecerão as seguintes condições :.
<blockquote style="margin-top:0;margin-bottom:0">a) nos cursos stricto sensu ou de especialização que fixem um número máximo de alunos por turma, são limitadas em 30% (trinta por cento) do total de vagas oferecidas; 
<br>b) nos cursos de pós-graduação lato sensu não haverá limites de vagas. Caso a estrutura do curso torne necessária a limitação do número de alunos será observado o disposto na alínea “a” deste parágrafo. 
</blockquote>
<b>Parágrafo oitavo</b> – Os bolsistas que forem reprovados no período letivo perderão o direito à bolsa de estudo, voltando a gozar do benefício quando lograrem aprovação no referido período. As disciplinas cursadas em regime de dependência serão de total responsabilidade do bolsista, arcando o mesmo com o seu custo. 
<br><b>Parágrafo nono</b> - Considera-se adquirido o direito daquele PROFESSOR que já esteja usufruindo bolsas de estudo em número superior ao definido nesta cláusula. 

<tr><td class=titulo>52. Autorização para desconto em folha de pagamento 
<tr><td class=campo style="text-align:justify">O desconto do professor em folha de pagamento somente poderá ser realizado mediante sua autorização, nos termos dos artigos 462 e 545 da CLT, quando os valores forem destinados ao custeio de prêmios de seguro, planos de saúde, mensalidades associativas ou outras que constem da sua expressa autorização, desde que não haja previsão expressa de desconto na presente norma coletiva. 
<br><b>Parágrafo único</b> – Encontra-se no SINPRO, à disposição da MANTENEDORA, cópia de autorização do PROFESSOR para o desconto da mensalidade associativa. 

<tr><td class=titulo>53. Estabilidade para portadores de doenças graves 
<tr><td class=campo style="text-align:justify">Fica assegurada, até alta médica, considerada como apto ao trabalho, ou eventual concessão de aposentadoria por invalidez, estabilidade no emprego aos PROFESSORES acometidos por doenças graves ou incuráveis e aos PROFESSORES portadores do vírus HIV que vierem a apresentar qualquer tipo de infecção ou doença oportunista, resultante da patologia de base. 
<br><b>Parágrafo único</b> – São consideradas doenças graves ou incuráveis, a tuberculose ativa, alienação mental, esclerose múltipla, neoplasia maligna, cegueira definitiva, hanseníase, cardiopatia grave, doença de Parkinson, paralisia irreversível e incapacitante, espondiloastrose anquilosante, neofropatia grave, estados do Mal de Paget (osteíte deformante) e contaminação grave por radiação. 

<tr><td class=titulo>54. Garantias ao professor com seqüelas ocasionadas por doenças profissionais ou acidente de trabalho 
<tr><td class=campo style="text-align:justify">Será garantida ao PROFESSOR acidentado no trabalho ou acometido por doença profissional a permanência na empresa em função compatível com o seu estado físico, sem prejuízo na remuneração antes percebida, desde que, após o acidente ou comprovação da aquisição de doença profissional, apresente, cumulativamente, redução da capacidade laboral, atestada pelo órgão oficial e que se tenha tornado incapaz de exercer a função que anteriormente desempenhava, obrigado, porém, o PROFESSOR nessa situação a participar dos processos de readaptação e reabilitação profissional. 
<br><b>Parágrafo único</b> – O período de estabilidade do PROFESSOR que se encontre participando dos processos de readaptação e reabilitação profissional será o previsto em lei. 

<tr><td class=titulo>55. Multa por descumprimento da Convenção
<tr><td class=campo style="text-align:justify">O descumprimento desta Convenção obrigará a MANTENEDORA ao pagamento de multa correspondente a 1% (um por cento) do salário do PROFESSOR, para cada uma das cláusulas não-cumpridas, acrescidas de juros, a cada PROFESSOR prejudicado. 
<br><b>Parágrafo único</b> – A MANTENEDORA está desobrigada de arcar com a multa prevista no caput, caso a cláusula descumprida já estabeleça uma multa pelo seu não–cumprimento. 

<tr><td class=titulo>56. Disposições transitórias 
<tr><td class=campo style="text-align:justify">Fica mantida a “Comissão de Aprimoramento das Relações de Trabalho”, composta de forma paritária, por quatro membros de cada uma das categorias profissional e econômica, indicados pela FEPESP e pelo SEMESP e/ou SEMESP/SJ RIO PRETO, com o objetivo de apresentar, até 30 de novembro de 2008, proposta de regulamentação dos seguintes temas 
<blockquote style="margin-top:0;margin-bottom:0">a) relações de trabalho envolvendo aplicações de novas tecnologias, cursos semipresenciais e cursos modulares e seqüenciais; 
<br>b) planos de carreira das Instituições privadas de ensino; 
</blockquote>
<b>Parágrafo primeiro</b> – A primeira reunião da “Comissão de Aprimoramento das Relações de Trabalho” será realizada às 10 horas do dia 27 de maio de 2008, na sede da FEPESP, em São Paulo, quando ocorrerá a aprovação do regimento de funcionamento. 
<br><b>Parágrafo segundo</b> – Os estudos, relatórios e deliberações da “Comissão de Aprimoramento das Relações do Trabalho”, serão submetidos às deliberações das Assembléias convocadas pelas respectivas entidades sindicais e, uma vez aprovadas, incluídas na presente Convenção, a partir da próxima data base, em 1º de março de 2009. 

<tr><td class=campo style="text-align:justify">E por estarem justos e acertados, assinam a presente Convenção Coletiva de Trabalho, a qual será depositada na Delegacia Regional do Trabalho de São Paulo, nos termos do artigo 614 e parágrafos, para fins de arquivo, de modo a surtir, de imediato, os seus efeitos legais. 

<tr><td class=campo style="text-align:justify">São Paulo, 30 de maio de 2008 

<br>
<pre>
<br>Hermes Ferreira Figueiredo                  Celso Napolitano 
<br>Presidente do SEMESP                        Presidente da FEPESP 
<br>
<br>Hélder Abud Paranhos                        Luiz Antonio Barbagli 
<br>Presidente do SINPRO Sorocaba               Presidente do SINPRO – SÃO PAULO 
<br>
<br>Rubens Gonçalves Aniz                       Marco Aurélio Arruda Aranha 
<br>Presidente do SINPRO – Osasco               Presidente do SINPRO – Salto, Indaiatuba e Itu 
<br>
<br>Martinho Condini                            Reginaldo Alberto Meloni 
<br>Presidente do SINPRO – Jundiaí              Presidente do SINPRO – Campinas 
<br>
<br>Aloísio Alves da Silva                      Rubens Gabriel Abdal 
<br>Presidente do SINPRO – ABC                  Presidente do SINPRO – Valinhos e Vinhedo 
<br>
<br>Ildefonso Paz Dias                          Andréa Luciana Harada Sousa 
<br>Presidente do SINPRO – Santos               Presidente do SINPRO – Guarulhos 
<br>
<br>Márcio Silva Sampaio Lopes                  Samuel Cristiano Fávero 
<br>Presidente do SINPRO Mogi Guaçu e Itapira   Presidente do SINPRO - Jau
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

