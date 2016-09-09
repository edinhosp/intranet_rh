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
<title>Convenção Coletiva 2008 - Auxiliares</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->

<!-- <b>AUXILIARES</b> -->
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>

<tr><td class=titulo align="center">CONVENÇÃO COLETIVA DE TRABALHO 2008/2010</td></tr>
<tr><td class=titulo align="center">ensino superior</td></tr>
<tr><td class=titulo align="center">Entidade Sindical Profissional – Auxiliares de Administração Escolar</td></tr>
<tr><td class=titulo align="center">Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de S. Paulo SEMESP</td></tr>

<tr><td class=campo style="text-align:justify">Entre as partes, de um lado, a FETEE - Federação dos Trabalhadores em Estabelecimento de Ensino do Estado de São Paulo, CNPJ nº 62197082/0001-63, <b>Sindicato dos Auxiliares de Administração Escolar do ABC – SAAE ABC</b>, CNPJ nº 69.116.069/0001-81; Sindicato dos Professores e Auxiliares Administrativos de Araçatuba e Região (Araçatuba e Birigui), CNPJ nº 00.376.088/0001-40; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Araraquara, CNPJ nº 66.994.393/0001-04; Sindicato dos Professores e Auxiliares de Administração Escolar de Bragança Paulista, CNPJ nº 61.699.666/0001-74; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Capivari; CNPJ nº 04.546.257/0001-02; Sindicato dos Professores e Trabalhadores em Educação de Dracena e Região (Junqueirópolis, Monte Castelo, Nova Guataporanga, Ouro Verde, Panorama, Paulicéia, Santa Mercedes, São João do Pau D´Alho, Tupi Paulista), CNPJ nº 64.615.461/0001-51; Sindicato dos Professores e Auxiliares Administrativos de Fernandópolis (Auriflama, Estrela D´Oeste, General Salgado, Ilha Solteira, Nhandeara, Pereira Barreto, Santa Fé do Sul, Urânia), CNPJ nº 63.893.838/0001-71; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Franca, CNPJ nº 60.239.845/0001-66; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Guaratinguetá, CNPJ nº 06.343.424/0001-35; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Itatiba; CNPJ nº 58.387.358/0001-07; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Jaguariúna (Pedreira, Santo Antônio da Posse, Holambra, Arthur Nogueira, Estiva Gerbi, Engenheiro Coelho, Conchal, Cosmópolis e Paulínia) CNPJ nº 06.368.966/001-62; Sindicato dos Professores e Auxiliares Administrativos de Jales, CNPJ nº 63.891.998/0001-81; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Leme, Pirassununga, Porto Ferreira e Descalvado, CNPJ nº 08.369.686/0001-02; Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Lins, CNPJ nº 51.520.187/0001-95; Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Lorena, CNPJ nº 65.042.038/0001-72; Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Marília, CNPJ nº 51.513.679/0001-53; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação Pindamonhangaba, CNPJ nº 07.192.010/0001-15; Sindicato dos Auxiliares de Administração Escolar de Piracicaba, CNPJ nº 56.979.545/0001-46; Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Presidente Prudente, CNPJ nº 53.301.305/0001-08; Sindicato dos Professores e Auxiliares de Administração Escolar de Ribeirão Preto, CNPJ nº 56.891.377/0001-32; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Rio Claro, CNPJ nº 55.360.846/0001-24; Sindicato dos Auxiliares de Administração Escolar de Santos, CNPJ nº 71.547.715/0001-07; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de São Carlos, CNPJ nº 06.266.000/0001-14; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de São João da Boa Vista, CNPJ nº 06.967.961/0001-56; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Sumaré, Hortolândia e Nova Odessa, CNPJ nº 07.493.086/0001-80; Sindicato dos Trabalhadores em Estabelecimento de Ensino e Educação de Taubaté, CNPJ nº 07.288.958/0001-79; Sindicato dos Professores e Auxiliares de Administração Escolar de Votuporanga, CNPJ nº 59.857.755/0001-50, entidades com bases territoriais e representatividades fixadas nas respectivas Cartas Sindicais e no que estabelece o inciso I do artigo 8º da Constituição Federal e de outro, o Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de São Paulo - SEMESP, CNPJ nº 49.343.874/0001-30, com representatividade fixada em sua Carta Sindical, ao final assinados por seus representantes legais, devidamente autorizados pelas competentes Assembléias Gerais das respectivas categorias, fica estabelecida, nos termos do artigo 611 e seguintes da Consolidação das Leis do Trabalho e do artigo 8º, inciso VI da Constituição Federal, a presente CONVENÇÃO COLETIVA DE TRABALHO.</td></tr>

<tr><td class=titulo>1. Abrangência</td></tr>
<tr><td class=campo style="text-align:justify">Esta Convenção Coletiva de Trabalho abrange a categoria profissional “AUXILIARES DE ADMINISTRAÇÃO ESCOLAR” (empregados em estabelecimentos de ensino), do 1º grupo – Trabalhadores em Estabelecimentos de Ensino – do plano da Confederação Nacional dos Trabalhadores em Estabelecimentos de Educação e Cultura, em dia com as suas obrigações estatutárias e das deliberações da Assembléia, doravante designados como “AUXILIARES” e a categoria econômica “estabelecimentos de ensino superior do Estado de São Paulo”, integrante do 1º grupo – Estabelecimentos de Ensino – do plano da Confederação Nacional de Educação e Cultura, representados pelo Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de São Paulo, doravante designados como “MANTENEDORAS”. 
<br><b>Parágrafo único</b> – A categoria profissional dos AUXILIARES DE ADMINISTRAÇÃO ESCOLAR abrange todos aqueles que, sob qualquer título ou denominação, exercem atividades não docentes nos estabelecimentos particulares de ensino superior, consoante a representação contida em sua Carta Sindical.</td></tr>

<tr><td class=titulo>2. Duração</td></tr>
<tr><td class=campo style="text-align:justify">Esta Convenção Coletiva de Trabalho tem vigência a partir da data de assinatura da presente e encerra-se em 28 de fevereiro de 2010.
<br><b>Parágrafo único</b> – As cláusulas poderão ser reexaminadas na próxima data base, em 1º de março de 2009, em virtude de problemas surgidos na sua aplicação ou do surgimento de normas legais a elas pertinentes, ou em decorrência de aprovação das propostas apresentadas pela Comissão Permanente de Negociação, prevista na cláusula 39 da presente Convenção.</td></tr>

<tr><td class=titulo>3. Reajuste salarial em 2008</td></tr>
<tr><td class=campo style="text-align:justify">I. Em 1º de dezembro de 2008, as MANTENEDORAS deverão aplicar o reajuste de 5,5% (cinco e meio por cento), sobre os salários devidos em 1º de fevereiro de 2008.</td></tr>
<tr><td class=campo style="text-align:justify">II. Considerando a data da assinatura da presente convenção coletiva, exclusivamente nos salários de dezembro de 2008, janeiro e fevereiro de 2009, a titulo de recomposição salarial, será acrescido o valor correspondente a 4,66% (quatro vírgula sessenta e seis por cento) do salário do mês fevereiro de 2008.</td></tr>
<tr><td class=campo style="text-align:justify">III. Considerando a data da assinatura da presente convenção coletiva, exclusivamente nos salários de março, abril, maio, junho e julho de 2009, a titulo de recomposição salarial, será acrescido o valor correspondente a 5,5% (cinco e meio por cento) do salário do mês fevereiro de 2008. A partir do mês de agosto de 2009, o valor correspondente a 5,5% (cinco e meio por cento) deixará de ser pago.
<br><b>Parágrafo primeiro</b> – As recomposições referidas nos incisos II e III desta cláusula, deverão ser registradas no comprovante de pagamento como rubrica própria e em destaque.
<br><b>Parágrafo segundo</b> – Fica estabelecido que o salário de 1º de dezembro de 2008, sem o valor correspondente à recomposição salarial, reajustado pelo índice definido nesta cláusula, servirá como base de cálculo para a data base de 1º de março de 2009.
<br><b>Parágrafo terceiro</b> - Para as Mantenedoras que concederam percentuais inferiores ao estabelecido na presente norma, referente aos meses de abril a novembro de 2008, as diferenças deverão ser pagas nas mesmas datas definidas no caput deste artigo, a título de recomposição salarial, observado o previsto no parágrafo primeiro,
<br><b>Parágrafo quarto</b> – Para as Mantenedoras que concederam antecipações salariais nos mesmos percentuais previstos na presente norma, no período de março a novembro de 2008, ficam isentas do pagamento referido nos incisos II e III do caput.</td></tr>

<tr><td class=titulo>4. Reajuste salarial em 1º de março de 2009</td></tr>
<tr><td class=campo style="text-align:justify">Em 1º de março de 2009, as MANTENEDORAS deverão aplicar sobre os salários devidos em 1º de dezembro de 2008, o percentual definido pela média aritmética dos índices inflacionários do período compreendido entre 1º de março de 2008 e 28 de fevereiro de 2009, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV), composto com 1,20% (um vírgula vinte por cento).
<br><b>Parágrafo primeiro</b> – O SEMESP e a ENTIDADE SINDICAL PROFISSIONAL comprometem-se a divulgar, em comunicado conjunto, até 20 de março de 2009, o percentual de reajuste salarial calculado pela fórmula definida no caput.
<br><b>Parágrafo segundo</b> – A base de cálculo para a data-base de 1º de março de 2010 será constituída pelos salários devidos em 1º de novembro de 2008, reajustados em 2009 pela fórmula definida no caput.</td></tr>

<tr><td class=titulo>5. Compensações salariais</td></tr>
<tr><td class=campo style="text-align:justify">No ano de 2008 será permitida a compensação de eventuais antecipações salariais concedidas no período compreendido entre 1º de março de 2008 a 1º de dezembro de 2008, substituindo as recomposições salariais previstas na cláusula 3. Relativamente à data-base de março de 2009 será permitida a compensação de eventuais antecipações salariais concedidas no período compreendido entre 1º de dezembro de 2008 e 28 de fevereiro de 2009.
<br><b>Parágrafo único</b> – Não serão permitidos, em ambos os casos, a compensação daquelas antecipações salariais que decorrerem de promoções, transferências, ascensão em plano de carreira e os reajustes concedidos com cláusula expressa de não–compensação.</td></tr>

<tr><td class=titulo>6. Salário do auxiliar ingressante na mantenedora</td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA não poderá contratar nenhum AUXILIAR por salário inferior ao limite salarial mínimo dos AUXILIARES mais antigos que possuam o mesmo grau de qualificação ou titulação de quem está sendo contratado, respeitado o quadro de carreira da MANTENEDORA.
<br><b>Parágrafo único</b> – Ao AUXILIAR admitido após 1º de dezembro de 2008 e após 1º de março de 2009, serão concedidos os mesmos percentuais de reajustes e aumentos salariais estabelecidos nas cláusulas 3 e 4, respectivamente, desta norma coletiva.</td></tr>

<tr><td class=titulo>7. Prazo e forma de pagamento dos salários</td></tr>
<tr><td class=campo style="text-align:justify">Os salários deverão ser pagos, no máximo, até o 5º dia útil do mês subseqüente ao trabalhado.
<br><b>Parágrafo primeiro</b> – O não pagamento dos salários no prazo obriga a MANTENEDORA a pagar multa diária, em favor do AUXILIAR, no valor de 1/30 (um trinta avos) de seu salário mensal.
<br><b>Parágrafo segundo</b> – As MANTENEDORAS que não efetuarem o pagamento dos salários em moeda corrente deverão proporcionar aos AUXILIARES tempo hábil para o recebimento no banco ou no posto bancário, excluindo-se o horário de refeição.
<br><b>Parágrafo terceiro</b> – As MANTENEDORAS que eventualmente alegarem impossibilidade de cumprimento do prazo estabelecido no parágrafo anterior, poderão requerer ao Foro Conciliatório outra data de pagamento de salários, desde que não ultrapasse o décimo dia do mês, ficando sujeitas às decisões adotadas no mesmo.</td></tr>

<tr><td class=titulo>8. Comprovantes de pagamento</td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA deverá fornecer ao AUXILIAR, mensalmente, comprovante de pagamento, devendo estar discriminados, quando for o caso:
<blockquote style="margin-top:0;margin-bottom:0">a) identificação da MANTENEDORA e do Estabelecimento de Ensino;
<br>b) identificação do AUXILIAR;
<br>c) denominação da função, se houver faixas salariais diferenciadas;
<br>d) carga horária mensal;
<br>e) outros eventuais adicionais;
<br>f) descanso semanal remunerado;
<br>g) horas extras realizadas;
<br>h) valor do recolhimento do FGTS;
<br>i) desconto previdenciário; e
<br>j) outros descontos.
</blockquote></td></tr>

<tr><td class=titulo>9. Adicional noturno</td></tr>
<tr><td class=campo style="text-align:justify">O adicional noturno deve ser pago nas atividades realizadas após as 22 horas e corresponde a 25% (vinte e cinco por cento) do valor das horas trabalhadas.</td></tr>

<tr><td class=titulo>10. Horas extras</td></tr>
<tr><td class=campo style="text-align:justify">Considera-se atividade extra todo trabalho desenvolvido em horário diferente daquele habitualmente realizado na semana. As três primeiras horas extras semanais devem ser pagas com o adicional de 50% (cinqüenta por cento) e as seguintes, com o adicional de 100% (cem por cento).
<br><b>Parágrafo primeiro</b> – Caso a MANTENEDORA implante o sistema de Banco de Horas deverá ser observado o disposto na cláusula própria que regula a matéria, integrante da presente norma coletiva.
<br><b>Parágrafo segundo</b> – Exceto nas hipóteses de necessidade comprovada, quando deverá ser produzido acordo expresso entre o AUXILIAR e a MANTENEDORA, é vedado, a esta, exigir, daquele, a realização de trabalhos ou qualquer outra atividade aos domingos e feriados. Havendo o acordo e não sendo concedida folga compensatória, fica assegurada a remuneração em dobro do trabalho realizado em tais dias, sem prejuízo do pagamento do repouso semanal remunerado.</td></tr>

<tr><td class=titulo>11. Adicional por atividades em outros municípios</td></tr>
<tr><td class=campo style="text-align:justify">Quando o AUXILIAR desenvolver suas atividades, em caráter eventual, a serviço da mesma MANTENEDORA, em município diferente daquele onde foi contratado e onde ocorre a prestação habitual do trabalho, deverá receber um adicional de 25% (vinte e cinco por cento) sobre o total de sua remuneração no novo município. Quando o AUXILIAR voltar a prestar serviços no município de origem, cessará a obrigação do pagamento deste adicional.
<br><b>Parágrafo primeiro</b> – Nos casos em que ocorrer a transferência definitiva do AUXILIAR, aceita livremente por este em documento firmado entre as partes, não haverá a incidência do adicional referido no “caput”, obrigando-se a MANTENEDORA a efetuar o pagamento de um único salário mensal integral, ao AUXILIAR, no ato de transferência, a título de ajuda de custo.
<br><b>Parágrafo segundo</b> – Fica assegurada a garantia de emprego pelo período de 6 (seis) meses ao AUXILIAR transferido de município, contados a partir do início do trabalho e/ou da efetivação da transferência.
<br><b>Parágrafo terceiro</b> – Caso a MANTENEDORA desenvolva atividade acadêmica em municípios considerados conurbanados, poderá solicitar isenção do pagamento do adicional determinado no caput, desde que encaminhe material comprobatório ao SEMESP, para análise e deliberação do Foro Conciliatório para Solução de Conflitos Coletivos, previsto na presente Convenção.</td></tr>

<tr><td class=titulo>12. Desconto de faltas</td></tr>
<tr><td class=campo style="text-align:justify">Na ocorrência de faltas não amparadas na legislação, a MANTENEDORA poderá descontar, no máximo, o número de horas em que o AUXILIAR esteve ausente e o DSR proporcional a essas horas, desde que a MANTENEDORA não tenha implantado o sistema de Banco de Horas conforme o disposto em cláusula própria da presente Convenção Coletiva de Trabalho.
<br><b>Parágrafo único</b> – É da competência e integral responsabilidade da MANTENEDORA estabelecer mecanismos de controle de faltas e de pontualidade do AUXILIAR, conforme a legislação vigente.</td></tr>

<tr><td class=titulo>13. Atestados médicos e abono de faltas</td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA é obrigada a aceitar atestados fornecidos por médicos ou dentistas conveniados ou credenciados pela entidade sindical profissional, SUS ou, ainda, por profissionais conveniados com a própria MANTENEDORA.
<br><b>Parágrafo único</b> – Também serão aceitos atestados que tenham sido convalidados pelas entidades sindicais de trabalhadores abrangidos por esta norma, pelos profissionais de saúde de departamento médico ou odontológico próprio ou conveniados às mesmas.</td></tr>

<tr><td class=titulo>14. Anotações na carteira de trabalho</td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA está obrigada a promover, em quarenta e oito horas, as anotações nas Carteiras de Trabalho de seus AUXILIARES, ressalvados eventuais prazos mais amplos permitidos por lei.
<br><b>Parágrafo único</b> – É obrigatória a anotação na CTPS das mudanças provocadas por ascensão em plano de carreira.</td></tr>

<tr><td class=titulo>15. Mudança de cargo ou função</td></tr>
<tr><td class=campo style="text-align:justify">O AUXILIAR não poderá ser transferido de um cargo ou função para outro, salvo com seu consentimento expresso e por escrito, sob pena de nulidade da referida transferência.</td></tr>

<tr><td class=titulo>16. Abono de faltas por casamento ou luto</td></tr>
<tr><td class=campo style="text-align:justify">Não serão descontadas, no curso de nove dias corridos, as faltas do AUXILIAR, por motivo de gala ou luto, este em decorrência de falecimento de pai, mãe, filho(a), cônjuge, companheiro(a) e dependente juridicamente reconhecido.
<br><b>Parágrafo único</b> – Em caso de falecimento de irmão(ã), sogro(a) e neto(a) os abonos ficarão reduzidos a três dias.</td></tr>

<tr><td class=titulo>17. Bolsas de estudo</td></tr>
<tr><td class=campo style="text-align:justify">Todo AUXILIAR tem direito a bolsas de estudo integrais, incluindo matrícula, para si, cônjuge, filhos ou dependentes legais, ambos entendidos como aqueles reconhecidos pela legislação do Imposto de Renda ou aqueles que estejam sob a guarda judicial do AUXILIAR e vivam sob sua dependência econômica, devidamente comprovada. Os filhos ou dependentes legais do AUXILIAR poderão usufruir as bolsas de estudo integrais, sem qualquer ônus, desde que não tenham 25 (vinte e cinco) anos completos ou mais na data da efetivação da matrícula no curso superior.
As bolsas de estudo são válidas para cursos de graduação, pós-graduação ou seqüenciais existentes e administrados pela MANTENEDORA localizado(s) no mesmo município onde trabalha o AUXILIAR, observado o disposto nesta cláusula e parágrafos seguintes.
<br><b>Parágrafo primeiro</b> – O direito às bolsas de estudo passa a vigorar ao término do contrato de experiência, cuja duração não pode exceder de 90 (noventa) dias, conforme parágrafo único do artigo 445 da CLT.
<br><b>Parágrafo segundo</b> – A MANTENEDORA está obrigada a conceder até duas bolsas de estudo por AUXILIAR, na vigência desta norma, sendo que, nos cursos de graduação ou seqüenciais, não será possível que o bolsista conclua mais de um curso nesta condição.
<br><b>Parágrafo terceiro</b> – A utilização do benefício previsto nesta cláusula, caracterizada como doação por não impor qualquer contraprestação de serviços é transitória e não habitual e, por isso, não possui caráter remuneratório e nem se vincula, para nenhum efeito, ao salário ou remuneração percebida pelo AUXILIAR, nos termos da Lei 10.243, de 19 de junho de 2001 e visa a capacitação dos beneficiários.
<br><b>Parágrafo quarto</b> – As bolsas de estudo serão mantidas quando o AUXILIAR estiver licenciado para tratamento de saúde ou em gozo de licença mediante anuência da MANTENEDORA, excetuado o disposto na cláusula da presente Convenção que trata sobre a Licença sem Remuneração.
<br><b>Parágrafo quinto</b> – No caso de falecimento do AUXILIAR, os dependentes que já se encontram estudando em estabelecimento de ensino superior da MANTENEDORA continuarão a gozar das bolsas de estudo até o final do curso, ressalvado o disposto no parágrafo 8º desta cláusula.
<br><b>Parágrafo sexto</b> – No caso de dispensa sem justa causa durante o período letivo, ficam garantidas ao AUXILIAR, até o final do período letivo, as bolsas de estudo já existentes.
<br><b>Parágrafo sétimo</b> – As bolsas de estudo integrais em cursos de pósgraduação ou especialização existentes e administrados pela MANTENEDORA são válidas exclusivamente para o AUXILIAR, em áreas correlatas àquelas em que o AUXILIAR exerce a função na MANTENEDORA e que visem à sua capacitação, respeitados os critérios de seleção exigidos para ingresso nos mesmos e obedecerão às seguintes condições:
<blockquote style="margin-top:0;margin-bottom:0">a) os cursos stricto sensu ou de especialização que fixem um número máximo de alunos por turma, são limitadas em 30% (trinta por cento) do total de vagas oferecidas;
<br>b) nos cursos de pós-graduação lato sensu não haverá limites de vagas. Caso a estrutura do curso torne necessária a limitação do número de alunos será observado o disposto na alínea a) deste parágrafo.
</blockquote>
    <b>Parágrafo oitavo</b> – Os bolsistas que forem reprovados no período letivo perderão o direito à bolsa de estudo, voltando a gozar do benefício quando lograrem aprovação no referido período. As disciplinas cursadas em regime de dependência serão de total responsabilidade do bolsista, arcando o mesmo com o seu custo.
<br><b>Parágrafo nono</b> – Considera-se adquirido o direito daquele AUXILIAR que já esteja usufruindo bolsas de estudo em número superior ao definido nesta cláusula.</td></tr>

<tr><td class=titulo>18. Irredutibilidade salarial</td></tr>
<tr><td class=campo style="text-align:justify">É proibida a redução da remuneração mensal ou de carga horária do AUXILIAR, exceto quando ocorrer iniciativa expressa do mesmo. Em qualquer hipótese, é obrigatória a concordância formal e recíproca, firmada por escrito.
<br><b>Parágrafo único</b> – Não havendo concordância recíproca, a parte que deu origem à redução prevista nesta cláusula arcará com a responsabilidade da rescisão contratual.</td></tr>

<tr><td class=titulo>19. Uniformes</td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA deverá fornecer gratuitamente dois uniformes por ano, quando o seu uso for exigido.</td></tr>

<tr><td class=titulo>20. Licença sem remuneração</td></tr>
<tr><td class=campo style="text-align:justify">O AUXILIAR, com mais de 5 (cinco) anos ininterruptos de serviço no estabelecimento ensino superior da MANTENEDORA, terá direito a licenciarse, sem direito à remuneração, por um período máximo de dois anos, não sendo este período de afastamento computado para contagem de tempo de serviço ou para qualquer outro efeito, inclusive legal.
<br><b>Parágrafo primeiro</b> – A licença ou sua prorrogação deverão ser comunicadas à MANTENEDORA com antecedência mínima de 90 (noventa) dias, devendo especificar as datas de início e término do afastamento. A licença só terá início a partir da data expressa no comunicado, mantendo-se, até aí, todas as vantagens contratuais. A intenção de retorno do AUXILIAR à atividade deverá ser comunicada à MANTENEDORA no mínimo 60 (sessenta) dias antes do término do afastamento.
<br><b>Parágrafo segundo</b> – O AUXILIAR que tenha ou exerça cargo de confiança deverá, junto com o comunicado de licença, solicitar seu desligamento do cargo a partir do início da licença.
<br><b>Parágrafo terceiro</b> – Considera-se demissionário o AUXILIAR que, ao término do afastamento, não retornar às atividades.</td></tr>

<tr><td class=titulo>21. Licença à auxiliar adotante</td></tr>
<tr><td class=campo style="text-align:justify">Nos termos da Lei nº 10.421, de 15 de abril de 2.002, será garantida licença maternidade às AUXILIARES que vierem a adotar ou obtiverem guarda judicial de crianças.</td></tr>

<tr><td class=titulo>22. Licença paternidade</td></tr>
<tr><td class=campo style="text-align:justify">A licença paternidade terá a duração de 5 dias.</td></tr>

<tr><td class=titulo>23. Garantia de emprego à gestante</td></tr>
<tr><td class=campo style="text-align:justify">Fica garantido emprego a AUXILIAR gestante desde o início da gravidez até sessenta dias após o término do afastamento legal. Em caso de dispensa, o aviso prévio começará a contar a partir do término do período de estabilidade.</td></tr>

<tr><td class=titulo>24. Creches</td></tr>
<tr><td class=campo style="text-align:justify">É obrigatória a instalação de local destinado à guarda de crianças até 12 meses, quando a unidade de ensino da MANTENEDORA mantiver contratadas, em jornada integral, pelo menos trinta funcionárias com idade superior a 16 anos. A manutenção da creche poderá ser substituída pelo pagamento do reembolso-creche, nos termos da legislação em vigor (CF, 7º, XXV, Artigo 389, parágrafo 1º da CLT e Portaria MTb nº 3296 de 03.09.86), ou ainda, a celebração de convênio com uma entidade reconhecidamente idônea.</td></tr>

<tr><td class=titulo>25. Garantias ao auxiliar em vias de aposentadoria</td></tr>
<tr><td class=campo style="text-align:justify">Fica assegurado ao AUXILIAR que, comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria por tempo de contribuição ou da aposentadoria por idade, a garantia de emprego durante o período que faltar até a aquisição do direito.
<br><b>Parágrafo primeiro</b> – A garantia de emprego é devida ao AUXILIAR que esteja contratado pela MANTENEDORA há pelo menos três anos.
<br><b>Parágrafo segundo</b> – A comprovação à MANTENEDORA deverá ser feita mediante a apresentação de documento que ateste o tempo de serviço. Este documento deverá ser emitido pelo INSS ou por pessoa credenciada junto ao órgão previdenciário. Se o AUXILIAR depender de documentação para realização da contagem, terá um prazo de 30 (trinta) dias, a contar da data prevista ou marcada para homologação da rescisão contratual.
<br><b>Parágrafo terceiro</b> – O contrato de trabalho do AUXILIAR só poderá ser rescindido por mútuo acordo homologado pelo sindicato ou por pedido de demissão.
<br><b>Parágrafo quarto</b> – Havendo acordo formal entre as partes, o AUXILIAR poderá exercer outra função compatível, durante o período em que estiver garantido pela estabilidade.
<br><b>Parágrafo quinto</b> – O aviso prévio, em caso de demissão sem justa causa, integra o período de estabilidade previsto nesta cláusula.
<br><b>Parágrafo sexto</b> – Enquanto não ocorrer a comprovação da documentação prevista nesta cláusula, o contrato de trabalho ficará suspenso. Caso o AUXILIAR não apresente a documentação até 30 (trinta) dias após a data prevista para homologação da rescisão, a demissão ocorrerá sem o pagamento de qualquer indenização adicional. Ocorrendo a comprovação da documentação, a rescisão contratual será cancelada e o AUXILIAR será reintegrado.</td></tr>

<tr><td class=titulo>26. Multa por atraso na homologação da rescisão contratual </td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA deve pagar as verbas devidas na rescisão contratual no dia seguinte ao término do aviso prévio, quando trabalhado, ou dez dias após o desligamento, quando houver dispensa do cumprimento de aviso prévio. A MANTENEDORA deve homologar a rescisão contratual até o 20º dia após o término do aviso prévio, quando trabalhado, ou trinta dias após o desligamento, quando houver dispensa do cumprimento de aviso prévio. O atraso na homologação obrigará a MANTENEDORA ao pagamento de multa, em favor do AUXILIAR, correspondente a um mês de sua remuneração. A partir do vigésimo dia de atraso, haverá ainda multa diária de 0,2% (dois décimos percentuais) do salário mensal. A MANTENEDORA está desobrigada de pagar a multa quando o atraso vier a ocorrer, comprovadamente, por motivos alheios à sua vontade.
<br><b>Parágrafo único</b> – A entidade sindical profissional está obrigada a fornecer comprovante de comparecimento sempre que a MANTENEDORA se apresentar para homologação das rescisões contratuais e comprovar a convocação do AUXILIAR.</td></tr>

<tr><td class=titulo>27. Demissão por justa causa</td></tr>
<tr><td class=campo style="text-align:justify">Quando houver demissão por justa causa, nos termos do art. 482, da CLT, a MANTENEDORA está obrigada a determinar na carta-aviso o motivo fático que deu origem à dispensa. Caso contrário, ficará descaracterizada a justa causa.</td></tr>

<tr><td class=titulo>28. Readmissão do auxiliar</td></tr>
<tr><td class=campo style="text-align:justify">O AUXILIAR que for readmitido para a mesma função até 12 (doze) meses após o seu desligamento ficará desobrigado de firmar contrato de experiência.</td></tr>

<tr><td class=titulo>29. Indenização por dispensa imotivada</td></tr>
<tr><td class=campo style="text-align:justify">O AUXILIAR demitido sem justa causa terá direito a uma indenização, além do aviso prévio legal de trinta dias e das indenizações previstas nesta Convenção, quando forem devidas, nas condições abaixo especificadas:
<blockquote style="margin-top:0;margin-bottom:0">a) 03 (três) dias para cada ano trabalhado na MANTENEDORA;
<br>b) aviso prévio adicional de quinze dias, caso o AUXILIAR tenha, no mínimo, cinqüenta anos de idade e que, à data do desligamento, conte com pelo menos um ano de serviço na MANTENEDORA.
</blockquote>
    <b>Parágrafo primeiro</b> – Não terá direito a indenização prevista na alínea “a” o AUXILIAR que tiver recebido, durante pelo menos um ano, pagamento mensal de adicional por tempo de serviço decorrente de plano de cargos e salários ou de anuênio, qüinqüênio ou equivalente, cujo valor corresponda a, no mínimo, 1% (um por cento) do valor do salário, por ano trabalhado. A MANTENEDORA deverá apresentar, no momento da homologação, documentos que comprovem o pagamento ao AUXILIAR do referido adicional por tempo de serviço.
<br><b>Parágrafo segundo</b> – Não terá direito à indenização assegurada na alínea “b” do caput, o AUXILIAR que, na data de admissão na MANTENEDORA, contar com mais de cinqüenta anos de idade.
<br><b>Parágrafo terceiro</b> – O pagamento das verbas indenizatórias previstas nesta cláusula não será cumulativo, cabendo ao AUXILIAR, no desligamento, o maior valor monetário entre os previstos nas alíneas “a” e “b” do caput.
<br><b>Parágrafo quarto</b> – Essas indenizações não contarão, para nenhum efeito, como tempo de serviço.</td></tr>

<tr><td class=titulo>30. Atestados de afastamento e salários</td></tr>
<tr><td class=campo style="text-align:justify">Sempre que solicitada, a MANTENEDORA deverá fornecer ao AUXILIARES atestado de afastamento e salário (AAS) previsto na legislação vigente.</td></tr>

<tr><td class=titulo>31. Férias</td></tr>
<tr><td class=campo style="text-align:justify">As férias dos AUXILIARES serão determinadas nos termos da legislação que rege a matéria, pela direção da MANTENEDORA, sendo admitida a compensação dos dias de férias concedidos antecipadamente, em período nunca inferior a 10 (dez) dias e nem mais que 2 (duas) vezes por ano.
<br><b>Parágrafo primeiro</b> – Fica assegurado aos AUXILIARES o pagamento, quando do início de suas férias, do salário correspondente às mesmas e do abono previsto no inciso XVII, artigo 7º, da Constituição Federal, no prazo previsto pelo artigo 145 da CLT, independentemente de solicitação pelos mesmos.
<br><b>Parágrafo segundo</b> – As férias, individuais ou coletivas, não poderão ter seu início coincidindo com domingos, feriados, dia de compensação do repouso semanal remunerado ou sábados, quando esses não forem dias normais de trabalho.</td></tr>

<tr><td class=titulo>32. Delegado representante</td></tr>
<tr><td class=campo style="text-align:justify">Em cada unidade que tenha mais de 50 AUXILIARES, a MANTENEDORA assegurará eleição de um Delegado Representante, que terá garantia de emprego e salários a partir da inscrição de sua candidatura até seis meses após o término de sua gestão, nos seguintes limites:
<blockquote style="margin-top:0;margin-bottom:0">a) Na unidade da MANTENEDORA que tenha até 100 (cem) AUXILIARES, será garantida a eleição de 01 (um) delegado representante;
<br>b) Na unidade da MANTENEDORA que tenha até mais de 200 (duzentos) AUXILIARES, será garantida a eleição de 02 (dois) delegados representantes;
</blockquote>
    <b>Parágrafo primeiro</b> – O mandato do Delegado Representante será de um ano.
<br><b>Parágrafo segundo</b> – A eleição do Delegado Representante será realizada pela entidade sindical na unidade de ensino da MANTENEDORA, por voto direto e secreto. É exigido quorum de 50% (cinqüenta por cento) mais um dos AUXILIARES da unidade de ensino da MANTENEDORA onde a eleição ocorrer.
<br><b>Parágrafo terceiro</b> – A entidade sindical comunicará a eleição à MANTENEDORA, com antecedência mínima de sete dias corridos. Nenhum candidato poderá ser demitido a partir da data da comunicação até o término da apuração.
<br><b>Parágrafo quarto</b> – É condição necessária que os candidatos sejam filiados a Entidade Sindical Profissional e que tenham, à data da eleição, pelo menos um ano de serviço na MANTENEDORA.</td></tr>

<tr><td class=titulo>33. Quadro de avisos</td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA deverá colocar à disposição da entidade sindical da categoria profissional quadro de avisos, em local visível, para fixação de comunicados de interesse da categoria, sendo proibida a divulgação de matéria político-partidária ou ofensiva a quem quer que seja.</td></tr>

<tr><td class=titulo>34. Assembléias sindicais</td></tr>
<tr><td class=campo style="text-align:justify">Todo AUXILIAR terá direito a abono de faltas para o comparecimento às assembléias da categoria.
<br><b>Parágrafo primeiro</b> – Na vigência desta Convenção, os abonos estão limitados, a dois sábados e mais dois dias úteis, quando a assembléia não for realizada no município em que o AUXILIAR trabalhe para a MANTENEDORA. Caso a Assembléia ocorra fora do município em que o AUXILIAR trabalhe para MANTENEDORA, os abonos estão limitados, a dois sábados e dois períodos. As duas assembléias realizadas durante os dias úteis deverão ocorrer em períodos distintos.
<br><b>Parágrafo segundo</b> – A entidade sindical deverá informar à MANTENEDORA, por escrito, com antecedência mínima de quinze dias corridos. Na comunicação deverão constar a data e o horário da assembléia.
<br><b>Parágrafo terceiro</b> – Os dirigentes sindicais não estão sujeitos ao limite previsto no parágrafo primeiro desta cláusula. As ausências decorrentes do comparecimento às assembléias de suas entidades serão abonadas mediante comunicação formal à MANTENEDORA.
<br><b>Parágrafo quarto</b> – A MANTENEDORA poderá exigir dos AUXILIARES e dos dirigentes sindicais atestado emitido pela entidade sindical profissional, que comprove o seu comparecimento à assembléia.</td></tr>

<tr><td class=titulo>35. Congressos, simpósios e equivalentes</td></tr>
<tr><td class=campo style="text-align:justify">Os abonos de falta para comparecimento a congressos, simpósios e equivalentes serão concedidos mediante aceitação por parte da MANTENEDORA, que deverá formalizar por escrito a dispensa do AUXILIAR.
<br><b>Parágrafo único</b> - A participação do AUXILIAR nos eventos descritos no “caput” não caracterizará atividade extraordinária.</td></tr>

<tr><td class=titulo>36. Congresso da entidade sindical profissional</td></tr>
<tr><td class=campo style="text-align:justify">Na vigência desta Convenção, a entidade sindical promoverá um evento de natureza política ou pedagógica (Congresso ou Jornada). A MANTENEDORA abonará as ausências de seus AUXILIARES que participarem do evento, nos seguintes limites:
<blockquote style="margin-top:0;margin-bottom:0">a) no estabelecimento de ensino superior que tenha até 49 AUXILIARES, será garantido, o abono a um AUXILIAR;
<br>b) no estabelecimento de ensino superior que tenha entre 50 e 99 AUXILIARES, será garantido, o abono a dois AUXILIARES;
<br>c) no estabelecimento de ensino superior que tenha mais de 100 AUXILIARES, será garantido, o abono a três AUXILIARES.
</blockquote>
Tais faltas, limitadas ao máximo de dois dias úteis além do sábado, serão abonadas mediante a apresentação de atestado de comparecimento fornecido pela entidade sindical. O AUXILIAR deverá repor as horas que, porventura, sejam necessárias para complementação da sua jornada de trabalho.</td></tr>

<tr><td class=titulo>37. Relação nominal</td></tr>
<tr><td class=campo style="text-align:justify">Obriga-se a MANTENEDORA a encaminhar para entidade representativa da categoria profissional, conforme Precedentes Normativos n.º 41 e 111, do Tribunal Superior do Trabalho, no prazo máximo de trinta dias contados da data do recolhimento da Contribuição Sindical, a relação nominal dos AUXILIARES que integram seu quadro de funcionários acompanhada do valor do salário mensal e das guias das contribuições sindical e assistencial.</td></tr>

<tr><td class=titulo>38. Foro conciliatório para solução de conflitos coletivos </td></tr>
<tr><td class=campo style="text-align:justify">Fica mantida a existência do Foro Conciliatório para Solução de Conflitos Coletivos, que tem como objetivo procurar resolver:
<br>I - divergências trabalhistas;
<br>II - incapacidade econômico-financeira da MANTENEDORA, no cumprimento de reajuste salarial e/ou de cláusulas previstas na presente convenção coletiva;
<br>III – alteração no prazo de pagamento de salários.
<br><b>Parágrafo primeiro</b> – Havendo dificuldade no cumprimento da cláusula de reajuste salarial ou diminuição nos percentuais de reajustes salariais estipulados nesta convenção coletiva ou definição de outro critério de reajuste salarial proposto pela MANTENEDORA, a solicitação da realização do Foro deverá ser formalizada por escrito e instruída com a documentação pertinente ao pedido.
<br><b>Parágrafo segundo</b> – Para efeito do que estabelece os incisos I, II e III deste artigo, a MANTENEDORA, ao solicitar o FORO, deve encaminhar os motivos do pedido de liberação do cumprimento da cláusula em questão, acompanhada da competente documentação comprobatória, para análise e decisão.
<br><b>Parágrafo terceiro</b> – O Foro será composto paritariamente, por três representantes do SEMESP e da ENTIDADE SINDICAL PROFISSIONAL. As reuniões deverão contar, também, com as partes em conflito que, se assim o desejarem, poderão delegar representantes para substituí-las e/ou serem assistidas por advogados, com poderes específicos para adotarem, em nome da Instituição, as decisões julgadas convenientes e necessárias.
<br><b>Parágrafo quarto</b> – O SEMESP e a ENTIDADE SINDICAL PROFISSIONAL deverão indicar os seus representantes no Foro num prazo de trinta dias a contar da assinatura desta Convenção.
<br><b>Parágrafo quinto</b> – Cada sessão do Foro será realizada no prazo máximo de quinze dias a contar da solicitação formal e obrigatória de qualquer uma das entidades que o compõem. A data, o local e o horário serão decididos pelas entidades sindicais envolvidas. O não comparecimento de qualquer uma das partes acarretará no encerramento imediato das negociações, bem como na aplicação na multa estabelecida no Parágrafo nono desta cláusula.
<br><b>Parágrafo sexto</b> – Nenhuma das partes envolvidas ingressará com ação na Justiça do Trabalho durante as negociações de entendimento.
<br><b>Parágrafo sétimo</b> – Na ausência de solução do conflito ou na hipótese de não comparecimento de qualquer uma das partes, a comissão responsável pelo Foro fornecerá certidão atestando o encerramento da negociação.
<br><b>Parágrafo oitavo</b> – Na hipótese de sucesso das negociações, a critério do Foro, a MANTENEDORA ficará desobrigada de arcar com a multa prevista no parágrafo 9 º (nono) desta cláusula.
<br><b>Parágrafo nono</b> – As decisões do Foro terão eficácia legal entre as partes acordantes. O descumprimento das decisões assumidas gerará multa a ser estabelecida no Foro, independentemente daquelas já estabelecidas nesta Convenção.
<br><b>Parágrafo dez</b> – A entidade sindical ou a MANTENEDORA que deixar de comparecer ao FORO, uma vez convocada, pagará uma multa de R$ 1.000,00 (hum mil reais), que reverterá em favor da parte presente.</td></tr>

<tr><td class=titulo>39. Comissão permanente de negociação</td></tr>
<tr><td class=campo style="text-align:justify">Fica mantida a Comissão Permanente de Negociação constituída de forma paritária, por três (3) representantes das entidades sindicais profissionais e econômica, com o objetivo de:
<blockquote style="margin-top:0;margin-bottom:0">a) fiscalizar o cumprimento das cláusulas vigentes;
<br>b) elucidar eventuais divergências de interpretação das cláusulas desta Convenção;
<br>c) discutir questões não-contempladas na norma coletiva;
<br>d) deliberar, no prazo máximo de trinta dias a contar da data da solicitação protocolizada no SEMESP, sobre modificação de pagamento da assistência médico-hospitalar, conforme os parágrafos 1º (primeiro) e 3º (terceiro) da cláusula relativa à matéria, constante desta norma coletiva;
<br>e) criar subsídios para a Comissão de Tratativas Salariais, através da elaboração de documentos para a definição das funções/atividades e o regime de trabalho dos AUXILIARES.
<br>f) criar critérios para a regionalização das negociações salariais referentes a 2010, bem como definir critérios diferenciados para elaboração do instrumento normativo destinado às entidades mantenedoras de Universidades, Centros Universitários, Faculdades, Institutos Superiores de Educação e Centros de Educação Tecnológicas.
</blockquote>
    <b>Parágrafo primeiro</b> – As entidades sindicais componentes da Comissão Permanente de Negociação indicarão seus representantes, no prazo máximo de trinta dias corridos, a contar da assinatura da presente Convenção.
<br><b>Parágrafo segundo</b> – A Comissão Permanente de Negociação deverá reunir-se mensalmente, em calendário elaborado de comum acordo entre as partes, alternadamente nas sedes das entidades sindicais que a compõem. Nos casos dispostos na letra “d” do caput, deverá haver convocação específica pela entidade sindical patronal.
<br><b>Parágrafo terceiro</b> – O não comparecimento da entidade sindical, profissional ou econômica, nas reuniões previstas no parágrafo 2º (segundo) da presente cláusula, implicará na multa de R$ 2.000,00 (dois mil reais) por reunião, a qual reverterá em benefício da entidade presente à mesma.</td></tr>

<tr><td class=titulo>40. Acordos internos</td></tr>
<tr><td class=campo style="text-align:justify">Ficam assegurados os direitos mais favoráveis decorrentes de acordos internos ou de acordos coletivos de trabalho celebrados entre a MANTENEDORA e a ENTIDADE SINDICAL PROFISSIONAL.</td></tr>

<tr><td class=titulo>41. Assistência médico-hospitalar</td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA está obrigada a assegurar, às suas expensas, assistência médico-hospitalar a todos os seus AUXILIARES, sendo-lhe facultada a escolha por plano de saúde, seguro-saúde ou convênios com empresas prestadoras de serviços médico-hospitalares. Poderá, ainda, prestar a referida assistência diretamente em se tratando de instituições que disponham de serviços de saúde e hospitais próprios ou conveniados. Qualquer que seja a opção feita, a assistência médico-hospitalar deve assegurar as condições e os requisitos mínimos que seguem relacionados:
<blockquote style="margin-top:0;margin-bottom:0">1. Abrangência – A assistência médico-hospitalar deve ser realizada no município onde funciona o estabelecimento de ensino superior ou onde vive o AUXILIAR, a critério da MANTENEDORA. Em casos de emergência, deverá haver garantia de atendimento integral em qualquer localidade do Estado de São Paulo ou fixação, em contrato, de formas de reembolso.
<br>2. Coberturas mínimas:
<blockquote style="margin-top:0;margin-bottom:0">2.1 Quarto para quatro pacientes, no máximo.
<br>2.2 Consultas.
<br>2.3 Prazo de internação de 365 dias por ano (comum e UTI/CTI)
<br>2.4 Parto, independentemente do estado gravídico.
<br>2.5 Moléstias infecto-contagiosas que exijam internação.
<br>2.6 Exames laboratoriais, ambulatoriais e hospitalares.
</blockquote>
    3. Carência – Não haverá carência na prestação dos serviços médicos e laboratoriais.
<br>4. Auxiliar ingressante – Não haverá carência para o AUXILIAR ingressante, independentemente do mês em que for contratado.
<br>5. Pagamento – A assistência médico-hospitalar será garantida nos termos desta Convenção, cabendo ao AUXILIAR, para usufruir dos benefícios da Lei nº 9656/98, o pagamento de 10% das mensalidades da referida assistência, respeitado o estabelecido no parágrafo 1º (primeiro) desta cláusula.
</blockquote>
    <b>Parágrafo primeiro</b> – Caso a assistência médico-hospitalar vigente na Instituição venha a sofrer reajuste em virtude de possíveis modificações estabelecidas em legislação que abranja o segmento – Lei 9.656, de 03 de junho de 1998 e MP 2.097-39, de 26 de abril de 2001 - ou que vierem a ser estabelecidas em lei, ou por mudança de empresa prestadora de serviço, a pedido do corpo técnico-administrativo da Instituição ou por quebra de contrato, unilateralmente, por parte da atual empresa prestadora de serviço, a MANTENEDORA continuará a contribuir com o valor mensal vigente até a data da modificação, devendo o AUXILIAR arcar com o valor excedente, que será descontado em folha e consignado no comprovante de pagamento, nos termos do art. 462, da CLT.
<br><b>Parágrafo segundo</b> – Caso ocorra mudança de empresa prestadora de serviço, por decisão unilateral da MANTENEDORA, com conseqüente reajuste no valor vigente, o AUXILIAR estará isento do pagamento do valor excedente, cabendo à MANTENEDORA prover integralmente a assistência médico-hospitalar, sem nenhum ônus para o AUXILIAR.
<br><b>Parágrafo terceiro</b> – Para efeito do disposto no Parágrafo primeiro desta cláusula, caberá à MANTENEDORA remeter a documentação comprobatória à Comissão Permanente de Negociação para a devida homologação.
<br><b>Parágrafo quarto</b> – Fica obrigado o AUXILIAR a optar pela prestação de assistência médico-hospitalar em uma única Instituição de ensino, quando mantiver mais de um vínculo empregatício como AUXILIAR no mesmo município ou municípios conurbanos. É necessário que o AUXILIAR se manifeste por escrito, com antecedência mínima de vinte dias, para que a MANTENEDORA possa proceder à suspensão dos serviços.
<br><b>Parágrafo quinto</b> – Mediante pagamento complementar e adesão facultativa, conforme o plano de atendimento médico-hospitalar e devidamente documentado, o AUXILIAR poderá optar pela ampliação dos serviços de saúde garantidos nesta Convenção Coletiva ou estendê-los a seus dependentes.</td></tr>

<tr><td class=titulo>42. Salário do auxiliar admitido para substituição</td></tr>
<tr><td class=campo style="text-align:justify">Ao AUXILIAR admitido em substituição a outro desligado, qualquer que tenha sido o motivo do seu desligamento, será garantido, sempre, salário inicial igual ao menor salário na função existente no estabelecimento, curso, grau ou nível de ensino, respeitado o Plano de Cargos e Salários da MANTENEDORA, sem serem consideradas eventuais vantagens pessoais.</td></tr>

<tr><td class=titulo>43. Menor salário da categoria</td></tr>
<tr><td class=campo style="text-align:justify">Fica assegurado, a partir de 1º (primeiro) de dezembro de 2008, nos termos do inciso V, artigo 7º, da Constituição Federal, um menor salário da categoria equivalente a R$ 561,63 (quinhentos e sessenta e um reais e sessenta e três centavos) por jornada integral de trabalho (44 horas semanais).
<br>A partir de 1º (primeiro) de março de 2009, nos termos do inciso V, artigo 7º, da Constituição Federal, será assegurado um menor salário da categoria equivalente ao resultado apurado pela aplicação do reajuste previsto na cláusula 4 desta norma coletiva, sobre o valor do piso em 1º de novembro de 2008, por jornada integral de trabalho (44 horas semanais).</td></tr>

<tr><td class=titulo>44. Abono de ponto ao estudante</td></tr>
<tr><td class=campo style="text-align:justify">Fica assegurado o abono de faltas ao AUXILIAR estudante para prestação de exames escolares, condicionado à prévia comunicação à MANTENEDORA e comprovação posterior.</td></tr>

<tr><td class=titulo>45. Prorrogação da jornada do estudante</td></tr>
<tr><td class=campo style="text-align:justify">Fica permitida a prorrogação da jornada de trabalho ao AUXILIAR estudante, ressalvadas as hipóteses de conflito com horário de freqüência às aulas.</td></tr>

<tr><td class=titulo>46. Estabilidade provisória do alistando</td></tr>
<tr><td class=campo style="text-align:justify">É assegurada aos AUXILIARES em idade de prestação do serviço militar estabilidade provisória, desde o alistamento até sessenta dias após a baixa.</td></tr>

<tr><td class=titulo>47. Auxiliar afastado por doença</td></tr>
<tr><td class=campo style="text-align:justify">Ao AUXILIAR afastado do serviço por doença devidamente atestada pela Previdência Social ou por médico ou dentista credenciado pela MANTENEDORA, será garantido o emprego ou o salário, a partir da alta, por igual período ao do afastamento, limitado a 60 (sessenta) dias além do aviso prévio.</td></tr>

<tr><td class=titulo>48. Refeitórios</td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA que contar com mais de 300 (trezentos) AUXILIARES no mesmo estabelecimento de ensino superior por ela mantido e não conceder vale-refeição obriga-se a manter refeitório.
<br><b>Parágrafo único</b> – No estabelecimento de ensino superior da MANTENEDORA em que trabalhem menos de 300 (trezentos) AUXILIARES será obrigatório assegurar-lhes condições de conforto e higiene por ocasião das refeições.</td></tr>

<tr><td class=titulo>49. Cesta básica</td></tr>
<tr><td class=campo style="text-align:justify">Fica assegurada aos AUXILIARES que percebam, até 4 (quatro) vezes o piso salarial da categoria, em jornada integral de 44 (quarenta e quatro) horas semanais, ou percebam, em jornada inferior, remuneração proporcionalmente igual ou inferior ao limite fixado nesta cláusula, a concessão de uma cesta básica mensal de 26 kg, composta, no mínimo, dos seguintes produtos não perecíveis:
<div align="center"><table width=350 border=0>
<tr><td class=campo>Arroz            </td><td class=campo>Óleo                </td><td class=campo>Macarrão </td></tr>
<tr><td class=campo>Feijão           </td><td class=campo>Café                </td><td class=campo>Sal </td></tr>
<tr><td class=campo>Farinha de Trigo </td><td class=campo>Farinha de Mandioca </td><td class=campo>Farinha de Milho </td></tr>
<tr><td class=campo>Açúcar           </td><td class=campo>Biscoito            </td><td class=campo>Purê de Tomate </td></tr>
<tr><td class=campo>Tempero          </td><td class=campo>Achocolatado        </td><td class=campo>Leite em Pó </td></tr>
<tr><td class=campo>Fubá             </td><td class=campo>Sardinha em Lata    </td><td class=campo>Sopão </td></tr>
</table></div>

<br><b>Parágrafo primeiro</b> – As MANTENEDORAS que já concedem vale-refeição, conforme o determinado pelo PAT, estão desobrigadas do fornecimento de cesta básica.
<br><b>Parágrafo segundo</b> – Fica assegurada a concessão de cesta básica durante as férias, licença maternidade e licença doença, bem como será garantido ao AUXILIAR demitido sem justa causa, na vigência da presente Convenção, a cesta básica referente ao período de aviso prévio, ainda que indenizado.</td></tr>

<tr><td class=titulo>50. Compensação semanal da jornada de trabalho
<tr><td class=campo style="text-align:justify">Fica permitida a compensação semanal da jornada de trabalho, nos termos da legislação que rege a matéria e obedecido o seguinte critério:
<blockquote style="margin-top:0;margin-bottom:0">a) mediante ciência, através do calendário anual a ser publicado pela MANTENEDORA, os AUXILIARES serão dispensados do cumprimento de sua jornada de trabalho em dias ali previstos, compensando-se as horas não trabalhadas com horas de trabalho complementares.
</blockquote></td></tr>

<tr><td class=titulo>51. Banco de horas</td></tr>
<tr><td class=campo style="text-align:justify">Nos termos da Lei nº 9.601, de 21 de janeiro de 1998, fica celebrado o Banco de Horas entre os AUXILIARES e as MANTENEDORAS, conforme o modelo descrito no parágrafo terceiro desta cláusula.
<br><b>Parágrafo primeiro</b> – As MANTENEDORAS que desejarem implantar o Banco de Horas, conforme o disposto no caput, deverão comunicar à entidade representativa da categoria profissional a implantação do mesmo, sob pena de não o fazendo não ter validade a aplicabilidade do Banco de Horas.
<br><b>Parágrafo segundo</b> – Caso a MANTENEDORA queira fazer alterações no Banco de Horas devido as suas peculiaridades, os critérios, detalhes, prazos e datas de implantação serão objeto de Acordo Coletivo de Trabalho específico, firmado entre a MANTENEDORA e seus AUXILIARES, com a participação da entidade sindical representativa da categoria profissional, na forma da legislação em vigor.
<br><b>Parágrafo terceiro</b> – O banco de horas deverá observar o seguinte modelo:
<div align="center"><table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=600>

<tr><td class=titulo align="center">ACORDO COLETIVO DE TRABALHO PARA A INSTITUIÇÃO DE BANCO DE HORAS</td></tr>
<tr><td class=campo style="text-align:justify"><b>Cláusula Primeira</b> – Nos termos da cláusula 50 da Convenção Coletiva de Trabalho 2008/10 firmada entre o SEMESP e a ENTIDADE SINDICAL PROFISSIONAL, fica estabelecido entre a (razão social da MANTENEDORA), neste ato representada pelo Sr. (nome e cargo que ocupa), e o SINDICATO DOS AUXILIARES DE ADMINISTRAÇÃO ESCOLAR de (base territorial), a criação do BANCO DE HORAS.</td></tr>
<tr><td class=campo style="text-align:justify"><b>Cláusula Segunda</b> – A partir de 1º de dezembro de 2008 fica instituído para a categoria dos AUXILIARES de Administração Escolar, o Sistema de Banco de Horas, com base na Lei 9.601/98, que deu nova redação ao § 2° do artigo 59 da Consolidação das Leis do Trabalho e a ele (art. 59) acrescentou o § 3°.
<br><b>§ 1º</b> Será formado um banco, proveniente das horas trabalhadas além da jornada normal diária, as quais serão compensadas nos termos do presente Acordo.
<br><b>§ 2º</b> A composição do Banco de Horas se dará mediante o acúmulo, apurado por meio de cartão de ponto, de horas credoras ou devedoras.
<br><b>§ 3º</b> As horas excedentes, a que se refere o parágrafo 2°, estarão limitadas a 2 (duas) horas diárias e 10 (dez) horas semanais, as quais serão acumuladas para futura compensação.
<br><b>§ 4º</b> Será permitido um saldo negativo de, no máximo, 20 horas a serem compensadas, conforme estabelecido nos parágrafos 6° a 12°.
<br><b>§ 5º</b> As horas que ultrapassarem o limite estabelecido no parágrafo 3° desta cláusula serão remuneradas como horas extras, em conformidade com o regulado em cláusula própria da Convenção Coletiva de Trabalho 2008.
<br><b>§ 6º</b> A compensação não poderá ocorrer nas Férias, Feriados e Descanso Semanal Remunerado.
<br><b>§ 7º</b> Sempre que houver interesse das partes em que haja a compensação, tal solicitação se dará com antecedência mínima de 48 (quarenta e oito) horas.
<br><b>§ 8º</b> A cada 120 (cento e vinte) dias serão realizados balanços para apuração do saldo de horas e planejamento da compensação, devendo tal saldo ser informado ao AUXILIAR. Havendo interesse entre as partes, o saldo existente poderá ser transferido, todo ou em parte, para o balanço do período seguinte. Poderá, ainda, o saldo apurado ser remunerado como hora extra, conforme o disposto na cláusula n. º 09 da Convenção Coletiva de Trabalho 2008/10.
<br><b>§ 9º</b> A apuração e compensação de saldo negativo obedecerá ao mesmo critério do parágrafo anterior.
<br><b>§ 10.</b> Os atrasos, saídas e faltas por motivo justificado e não previsto na legislação ou na CCT 2008/10, poderão ser compensados no Banco de Horas, limitando-se em uma ocorrência por semana.
<br><b>§ 11.</b> Os AUXILIARES contratados por prazo determinado, bem como aqueles que estão em período de experiência, não poderão valer-se do sistema de Banco de Horas.
<br><b>§ 12.</b> Nos casos de desligamento de AUXILIARES durante a vigência deste Acordo, obrigar-se-á a MANTENEDORA a pagar o adicional de Horas Extras sobre as horas não compensadas, calculadas sobre o valor da remuneração na data da rescisão. Na existência de horas a compensar (saldo negativo), conforme previsto nos parágrafos 6° e 9°, estas serão descontadas das verbas rescisórias.
<br><b>§ 13.</b> Qualquer divergência na aplicação deste Acordo deverá ser resolvida através da convocação do Foro para Solução de Conflitos Coletivos, conforme Cláusula específica da Convenção Coletiva de Trabalho.
<br><b>§ 14.</b> A renovação, alteração ou rescisão deste Acordo dependerá de acordo escrito dos representantes das partes, antes de expirado seu prazo de validade.
<br><b>§ 15.</b> O prazo de vigência do presente banco de horas é de 12 (doze) meses, encerrando-se em 28 de fevereiro de 2009.</td></tr>
<tr><td class=campo style="text-align:justify">(Data e local de assinatura, com identificação dos signatários)</td></tr>
</table> 
</div></td></tr>

<tr><td class=titulo>52. Autorização para desconto em folha de pagamento</td></tr>
<tr><td class=campo style="text-align:justify">O desconto do AUXILIAR em folha de pagamento somente poderá ser realizado, mediante sua autorização, nos termos dos artigos 462 e 545 da CLT, quando os valores forem destinados ao custeio de prêmios de seguro, planos de saúde, mensalidades associativas ou outras que constem da sua expressa autorização, desde que não haja previsão expressa de desconto na presente norma coletiva.
<br><b>Parágrafo único</b> – Encontra-se na entidade sindical profissional, à disposição da MANTENEDORA, cópia de autorização do AUXILIAR para o desconto da mensalidade associativa.</td></tr>

<tr><td class=titulo>53. Estabilidade para portadores de doenças graves</td></tr>
<tr><td class=campo style="text-align:justify">Fica assegurada, até alta médica, considerada como aptidão ao trabalho, ou eventual concessão de aposentadoria por invalidez, estabilidade no emprego aos AUXILIARES acometidos por doenças graves ou incuráveis e aos AUXILIARES portadores do vírus HIV que vierem a apresentar qualquer tipo de infecção ou doença oportunista, resultante da patologia de base.
<br><b>Parágrafo único</b> – São consideradas doenças graves ou incuráveis, a tuberculose ativa, alienação mental, esclerose múltipla, neoplasia maligna, cegueira definitiva, hanseníase, cardiopatia grave, doença de Parkinson, paralisia irreversível e incapacitante, espondiloartrose anquilosante, nefropatia grave, estados do Mal de Paget (osteíte deformante) e contaminação grave por radiação.</td></tr>

<tr><td class=titulo>54. Garantias ao auxiliar com sequelas e readaptação</td></tr>
<tr><td class=campo style="text-align:justify">Será garantida ao AUXILIAR acidentado no trabalho ou acometido por doença profissional, a permanência na MANTENEDORA em função compatível com seu estado físico, sem prejuízo da remuneração antes percebida, desde que após o acidente ou comprovação da aquisição de doença profissional apresente, cumulativamente, redução da capacidade laboral, atestada por órgão oficial e que se tenha tornado incapaz de exercer a função que anteriormente desempenhava, obrigado, porém, o AUXILIAR nessa situação a participar dos processos de readaptação e reabilitação profissionais.
<br><b>Parágrafo único</b> – O período de estabilidade do AUXILIAR que se encontra participando dos processos de readaptação e reabilitação profissionais será o previsto em lei.</td></tr>

<tr><td class=titulo>55. Competência das entidades sindicais signatárias</td></tr>
<tr><td class=campo style="text-align:justify">Fica estabelecida a legalidade das entidades sindicais signatárias para promover, perante a Justiça do Trabalho e o Foro em Geral, ações plúrimas em nome dos AUXILIARES em nome próprio, ou ainda, como parte interessada, em caso de descumprimento de qualquer cláusula avençada ou determinada nesta norma coletiva.</td></tr>

<tr><td class=titulo>56. Primeiros socorros</td></tr>
<tr><td class=campo style="text-align:justify">A MANTENEDORA obriga-se a manter materiais de primeiros socorros nos locais de trabalho e providenciar, por sua conta, a remoção do AUXILIAR acidentado/doente para o atendimento médico-hospitalar.</td></tr>

<tr><td class=titulo>57. Flexibilização da jornada de trabalho</td></tr>
<tr><td class=campo style="text-align:justify">Poderá ser flexibilizada a carga horária entre jornadas do AUXILIAR, quando no exercício concomitante de função docente e atividade administrativa, não havendo assim pagamento de salários nos intervalos, quando o AUXILIAR não tenha trabalhado nos mesmos.</td></tr>

<tr><td class=titulo>58. Multa por descumprimento da convenção</td></tr>
<tr><td class=campo style="text-align:justify">O descumprimento de cada cláusula desta Convenção obrigará a MANTENEDORA ao pagamento de multa correspondente a 5% (cinco por cento) do salário do AUXILIAR, acrescida de juros e correção monetária, a qual reverterá para a parte prejudicada.
<br><b>Parágrafo único</b> – A MANTENEDORA está desobrigada de arcar com o valor previsto nesta cláusula, caso o artigo da Convenção já estabeleça uma multa pelo não cumprimento da mesma.</td></tr>
<tr><td class=campo style="text-align:justify">Por estarem justos e acertados, assinam a presente Convenção Coletiva de Trabalho, a qual será depositada, para fins de arquivo, na Delegacia Regional do Trabalho e Emprego no Estado de São Paulo, nos termos do artigo 614, da Consolidação das Leis do Trabalho, de modo a surtir, de imediato, os seus efeitos legais.</td></tr>

<tr><td class=campo style="text-align:justify">São Paulo, 24 de novembro de 2008.
<br>
<pre>
Hermes Ferreira Figueiredo                              Geraldo Mugayar
Presidente do SEMESP                                    Presidente da FETEE - Federação dos Trab.
                                                        em Estab.de Ensino do Estado de São Paulo

Celso Soares Nogueira                                   Luiz Carlos Custódio
Sind. dos Aux. de Adm. Escolar do ABC – SAAE ABC        Presidente do Sind. dos Prof. e Aux. 
                                                        Administrativos de Araçatuba e Região

José Maria Gasparetto                                   Moacir Pereira
Sind. dos Trab. em Estab. de Ensino e Educação          Sind. dos Prof. e Aux. de Adm. Escolar
de Araraquara                                           de Bragança Paulista

Antonio Favarelli                                       Ronaldi Torelli
Sind. dos Trab. em Estab. de Ensino e Educação          Sind. dos Prof. e Trab. em Educação de 
de Capivari                                             Dracena e Região 

Cássio Antônio da Silva Tenani                          Regnério Terra
Sind. dos Prof. e Aux. Administrativos                  Sind. dos Trab. em Estab. de Ensino e 
de Fernandópolis                                        Educação de Franca
 
Reginaldo Costa                                         Remus Marin Stanc
Sind. dos Trab. em Estab. de Ensino e Educação de       Sind. dos Trab. em Estab. de Ensino e 
Guaratinguetá                                           Educação de Itatiba

Cássio Antônio da Silva Tenani                          Vera Lúcia Gorron
Sind. dos Prof. e Aux. Administrativos de Jales         Sind. dos Trab. em Estab. de Ensino e 
                                                        Educação de Leme, Pirassununga, Porto 
                                                        Ferreira e Descalvado

Ayrton Onofre da Silva                                  Hamilton Rosa Ferreira
Sind. dos Trab. em Estab.s de Ensino de Lins            Sind. dos Trab. em Estab.s de Ensino
                                                        de Lorena

José Roberto Marques de Castro                          Mário Joaquim Aredes Crescêncio
Sind. dos Trab. em Estab.s de Ensino de Marília         Sind. dos Trab. em Estab. de Ensino 
                                                        e Educação Pindamonhangaba

João Manoel dos Santos                                  Ademir Rodrigues
Sind. dos Aux. de Adm. Escolar de Piracicaba            Sind. dos Trab. em Estab.s de Ensino 
                                                        de Presidente Prudente

Antônio Dias de Novaes                                  Mara Lúcia Bito Legatzki
Sind. dos Prof. e Aux. de Adm. Escolar de               Sind. dos Trab. em Estab. de Ensino 
Ribeirão Preto                                          e Educação de Rio Claro

Márcio Campos                                           Maurício Carlos Ruggiero
Sind. dos Aux. de Adm. Escolar de Santos                Sind. dos Trab. em Estab. de Ensino 
                                                        e Educação de São Carlos

Francisco de Assis Carvalho Arten                       Sérgio Marcus Silva Franco
Sind. dos Trab. em Estab. de Ensino e Educação de       Sind. dos Trab. em Estab. de Ensino 
Sumaré, Hortolândia e Nova Odessa                       e Educação de São João da Boa Vista

Jeferson Campos                                         Armando Raphael D’ Avoglio
Sind. dos Trab. em Estab. de Ensino e Educação de       Sind. dos Prof. e Aux. de Adm. Escolar 
Taubaté                                                 de Votuporanga

Paulo Sérgio Silva Franco
Sind. dos Trab. em Estab. de Ensino e Educação de 
Jaguariúna e Região
</pre></td></tr>



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