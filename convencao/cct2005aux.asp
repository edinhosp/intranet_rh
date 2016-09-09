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
<title>Convenção Coletiva 2005 - Auxiliares</title>
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
<tr><td class=titulo align="center">CONVENÇÃO COLETIVA DE TRABALHO PARA 2005/2006</td></tr>
<tr><td class=titulo align="center">AUXILIARES DE ADMINISTRAÇÃO ESCOLAR
<tr><td class=titulo align="center">ENSINO SUPERIOR 
<tr><td class=campo style="text-align:justify">Entre as partes, de um lado, SINDICATO DOS ESTABELECIMENTOS DE ENSINO SUPERIOR NO ESTADO DE SÃO PAULO – SEMESP,entidade sindical de 1º grau, coordenadora e representativa dos estabelecimentos privados de ensino superior no Estado de São Paulo, com base territorial definida em sua Carta Sindical, inscrito no CNPJ sob nº 49343874/0001-30, Código Sindical nº Processo MTb 303127, com sede na rua Cipriano Barata nº 2431, Ipiranga, São Paulo, Capital, CEP 04205-002,com base territorial definida em sua Carta Sindical, em consonância com os incisos I e II, do artigo 8º, da Constituição Federal, representado por seu Presidente, Professor Hermes Ferreira Figueiredo, RG nº 2655493 - SSP/SP, CPF 04946158-34, devidamente autorizado para negociações e celebração de Convenção Coletiva de Trabalho, pela assembléia geral extraordinária realizada em 17 de março de 2005, conforme edital publicado no jornal Diário de São Paulo, edição de 1º de março de 2005, em cumprimento ao disposto na Instrução Normativa SRT/MTE nº 01, de 24 de março de 2004, publicada no DOU, Seção I, fls. 59 e 60, edição de 19 de abril de 2004, da Secretaria de Relações do Trabalho do Ministério do Trabalho e Emprego e de outro, FEDERAÇÃO DOS TRABALHADORES EM ESTABELECIMENTOS DE ENSINO DO ESTADO DE SÃO PAULO – FETEE/SP,registro sindical MTb nº 618670/48, CNPJ nº 062.197.082/0001-53, representada por seu Presidente, Professor Geraldo Mugayar, CPF 023779778-07, RG nº 1447287 – SSP/SP, também devidamente autorizada para negociações e assinatura de Convenção Coletiva de Trabalho, pela assembléia geral extraordinária realizada em 14 de dezembro de 2004, conforme editais publicados no Diário Oficial do Estado e em mais 34 (trinta e quatro) jornais de circulação estadual e regional, edição de 08 de dezembro de 2004, fica estabelecida, nos termos do artigo 611, § 2º, 613, 614 e seguintes, da Consolidação das Leis do Trabalho, do artigo 8º, VI, do artigo 7º, XXVI e artigo 5º, caput e inciso I, todos da Constituição Federal, a presente Convenção Coletiva de Trabalho:
<tr><td class=titulo>1. ABRANGÊNCIA
<tr><td class=campo style="text-align:justify">Esta Convenção Coletiva de Trabalho abrange a categoria profissional “AUXILIARES DE ADMINISTRAÇÃO ESCOLAR” (empregados em estabelecimentos de ensino), do 1º grupo – Trabalhadores em Estabelecimentos de Ensino – do plano da Confederação Nacional dos Trabalhadores em Estabelecimentos de Educação e Cultura, em dia com as suas obrigações estatutárias e das deliberações da Assembléia, doravante designados como “AUXILIARES” e a categoria econômica “estabelecimentos de ensino superior do Estado de São Paulo”, integrante do 1º grupo – Estabelecimentos de Ensino – do plano da Confederação Nacional de Educação e Cultura, representados pelo Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de São Paulo, doravante designados como “MANTENEDORAS”.
<br><b>Parágrafo único</b> – A categoria profissional dos <b>AUXILIARES</b> DE ADMINISTRAÇÃO ESCOLARabrange todos aqueles que, sob qualquer título ou denominação, exercem atividades não docentes nos estabelecimentos particulares de ensino superior.

<tr><td class=titulo>2. DURAÇÃO
<tr><td class=campo style="text-align:justify">Esta Convenção Coletiva de Trabalho terá a duração de dois anos, com vigência de 1º de março de 2005 a 28 de fevereiro de 2007.
<br><b>Parágrafo único</b> – As cláusulas constantes da presente norma poderão ser reexaminadas na próxima data-base, em virtude de problemas surgidos na sua aplicação ou do surgimento de normas legais a elas pertinentes, para as devidas adequações.

<tr><td class=titulo>3. REAJUSTE SALARIAL
<tr><td class=campo style="text-align:justify">A partir de 1º (primeiro) de maio de 2005 os salários dos <b>AUXILIARES</b> serão reajustados em 7,66 % ( sete virgula sessenta e seis por cento) incidentes sobre os salários devidos em 1º (primeiro) de fevereiro de 2005, reajustados conforme estabelece a Convenção Coletiva de 2004, observado o estabelecido na cláusula 4ª (quarta) da presente norma coletiva.
<br><b>Parágrafo primeiro</b> – Fica estabelecido que os salários de 1º (primeiro) de maio de 2005, reajustado pelo índice definido nesta cláusula, servirão como base de cálculo para a data-base de 1º (primeiro) de março de 2006.
<br><b>Parágrafo segundo</b> – Eventuais diferenças salariais resultantes da aplicação da presente norma coletiva, até a data de sua assinatura, deverão ser pagas até o dia 15 (quinze) de setembro de 2005, sem incidência da multa contratual.
<tr><td class=campo style="text-align:justify"><b>3.1. Reajuste Salarial em 1º de março de 2006
<tr><td class=campo style="text-align:justify">Em 1º (primeiro) de março de 2006, as MANTENEDORAS deverão aplicar sobre os salários devidos em 1º (primeiro) de maio de 2005, o percentual definido pela média aritmética dos índices inflacionários do período compreendido entre 1º (primeiro) de março de 2005 e 28 de fevereiro de 2006, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV).
<br><b>Parágrafo primeiro</b> – Se a média aritmética dos índices inflacionários definida no caput superar 9,99% (nove virgula noventa e nove por cento), as MANTENEDORAS deverão aplicar, em 1º de março de 2006, sobre os salários devidos em 1º de maio de 2005, o reajuste de 9,99% (nove virgula noventa e nove por cento). O SEMESP, a FETEE e os Sindicatos que representa, definirão, em processo de negociação salarial, até o prazo máximo de 30 de abril de 2006, a forma de pagamento da parcela excedente a 9,99%.
<br><b>Parágrafo segundo</b> – O SEMESP, a FETEE, e os Sindicatos que representa, comprometem-se a divulgar, em comunicado conjunto, até 20 de março de 2006, o percentual de reajuste salarial calculado pela fórmula definida no caput, bem como a forma de pagamento da parcela excedente a 9,99%, conforme estabelecido no parágrafo 1º (primeiro) desta cláusula.
<br><b>Parágrafo terceiro</b> – A base de cálculo para a data-base de 1º (primeiro) de março de 2007 será constituída pelos salários devidos em 1º (primeiro) de maio de 2005, reajustados em 2006 pela média aritmética dos índices inflacionários do período compreendido entre 1º de março de 2005 e 28 de fevereiro de 2006, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV).

<tr><td class=titulo>4. COMPENSAÇÕES SALARIAIS
<tr><td class=campo style="text-align:justify">Para 2005 será permitida a compensação de eventuais antecipações salariais concedidas no período de vigência da Convenção de 2004. Relativamente à convenção coletiva de 2006, será permitida a compensação de eventuais antecipações salariais concedidas no período de vigência da Convenção de 2005.

<tr><td class=titulo>5. SALÁRIO DO <b>AUXILIAR</b> INGRESSANTE NA MANTENEDORA
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> não poderá contratar nenhum <b>AUXILIAR</b> por salário inferior ao limite salarial mínimo dos <b>AUXILIARES</b> mais antigos que possuam o mesmo grau de qualificação ou titulação de quem está sendo contratado, respeitado o quadro de carreira da MANTENEDORA.
<br><b>Parágrafo único</b> - Ao <b>AUXILIAR</b> admitido após 1º de março de 2005 e após 1º de março de 2006, respectivamente, serão concedidos os mesmos percentuais de reajustes e aumentos salariais estabelecidos nesta norma coletiva.

<tr><td class=titulo>6. PRAZO E FORMA DE PAGAMENTO DOS SALÁRIOS
<tr><td class=campo style="text-align:justify">Os salários deverão ser pagos, no máximo, até o 5º dia útil do mês subseqüente ao trabalhado.
<br><b>Parágrafo primeiro</b> - O não pagamento dos salários no prazo obriga a <b>MANTENEDORA</b> a pagar multa diária, em favor do AUXILIAR, no valor de 1/30 (um trinta avos) de seu salário mensal.
<br><b>Parágrafo segundo</b> – As MANTENEDORAS que não efetuarem o pagamento dos salários em moeda corrente deverão proporcionar aos <b>AUXILIARES</b> tempo hábil para o recebimento no banco ou no posto bancário, excluindo-se o horário de refeição.
<br><b>Parágrafo terceiro</b> - As MANTENEDORAS que eventualmente alegarem impossibilidade de cumprimento do prazo estabelecido no parágrafo anterior, poderão requer ao Foro Conciliatório outra data de pagamento de salários, desde que não ultrapasse o décimo dia do mês, ficando sujeitas às decisões adotadas no mesmo.

<tr><td class=titulo>7. COMPROVANTES DE PAGAMENTO
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> deverá fornecer ao AUXILIAR, mensalmente, comprovante de pagamento, devendo estar discriminados, quando for o caso:
<blockquote style="margin-top:0;margin-bottom:0">a) identificação da <b>MANTENEDORA</b> e do Estabelecimento de Ensino;
<br>b) identificação do AUXILIAR;
<br>c) denominação da função, se houver faixas salariais diferenciadas;
<br>d) carga horária mensal;
<br>e) outros eventuais adicionais;
<br>f) descanso semanal remunerado;
<br>g) horas extras realizadas;
<br>h) valor do recolhimento do FGTS;
<br>i) desconto previdenciário;
<br>j) outros descontos.</blockquote>

<tr><td class=titulo>8. ADICIONAL NOTURNO
<tr><td class=campo style="text-align:justify">O adicional noturno deve ser pago nas atividades realizadas após as 22 horas e corresponde a 25% (vinte e cinco por cento) do valor das horas trabalhadas, a partir de maio de 2005.

<tr><td class=titulo>9. HORAS EXTRAS
<tr><td class=campo style="text-align:justify">Considera-se atividade extra todo trabalho desenvolvido em horário diferente daquele habitualmente realizado na semana. As três primeiras horas extras semanais devem ser pagas com o adicional de 50% (cinqüenta por cento) e as seguintes, com o adicional de 100% (cem por cento).
<br><b>Parágrafo primeiro</b> – Caso a <b>MANTENEDORA</b> implante o Banco de Horas deverá ser observado o disposto na cláusula que regula a matéria, integrante da presente norma coletiva.
<br><b>Parágrafo segundo</b> - Exceto nas hipóteses de necessidade comprovada, quando deverá ser produzido acordo expresso entre o <b>AUXILIAR</b> e a MANTENEDORA, é vedado, a esta, exigir, daquele, a realização de trabalhos ou qualquer outra atividade aos domingos e feriados. Havendo o acordo e não sendo concedida folga compensatória, fica assegurada a remuneração em dobro do trabalho realizado em tais dias, sem prejuízo do pagamento do repouso semanal remunerado.

<tr><td class=titulo>10. ADICIONAL POR ATIVIDADES EM OUTROS MUNICÍPIOS
<tr><td class=campo style="text-align:justify">Quando o <b>AUXILIAR</b> desenvolver suas atividades, em caráter eventual, a serviço da mesma MANTENEDORA, em município diferente daquele onde foi contratado e onde ocorre a prestação habitual do trabalho, deverá receber um adicional de 25% (vinte e cinco por cento) sobre o total de sua remuneração no novo município. Quando o <b>AUXILIAR</b> voltar a prestar serviços no município de origem, cessará a obrigação do pagamento deste adicional.
<br><b>Parágrafo primeiro</b> - Nos casos em que ocorrer a transferência definitiva do AUXILIAR, aceita livremente por este, em documento firmado entre as partes, não haverá a incidência do adicional referido no “caput”, obrigando-se a <b>MANTENEDORA</b> a efetuar o pagamento de um único salário mensal integral, ao AUXILIAR, no ato de transferência, a título de ajuda de custo.
<br><b>Parágrafo segundo</b> – Fica assegurada a garantia de emprego pelo período de 6 (seis) meses ao <b>AUXILIAR</b> transferido de município, contados a partir do início do trabalho e/ou da efetivação da transferência.
<br><b>Parágrafo terceiro</b> – Caso a <b>MANTENEDORA</b> desenvolva atividade acadêmica em municípios considerados conurbanados, poderá solicitar isenção do pagamento do adicional determinado no caput, desde que encaminhe material comprobatório ao SEMESP, para análise e deliberação do Foro Conciliatório para Solução de Conflitos Coletivos, previsto na presente Convenção.

<tr><td class=titulo>11. DESCONTO DE FALTAS
<tr><td class=campo style="text-align:justify">Na ocorrência de faltas não amparadas na legislação, a <b>MANTENEDORA</b> poderá descontar, no máximo, o número de horas em que o <b>AUXILIAR</b> esteve ausente e o DSR proporcional a essas horas, desde que a <b>MANTENEDORA</b> não tenha implantado o Banco de Horas conforme o disposto na presente Convenção Coletiva de Trabalho.
<br><b>Parágrafo único</b> - É da competência e integral responsabilidade da <b>MANTENEDORA</b> estabelecer mecanismos de controle de faltas e de pontualidade do <b>AUXILIAR</b> , conforme a legislação vigente.

<tr><td class=titulo>12. ATESTADOS MÉDICOS E ABONO DE FALTAS
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> é obrigada a aceitar atestados fornecidos por médicos ou dentistas conveniados ou credenciados pela entidade sindical profissional, SUS ou, ainda, por profissionais conveniados com a própria MANTENEDORA.
<br><b>Parágrafo único</b> - Também serão aceitos atestados que tenham sido convalidados pelas entidades sindicais de trabalhadores abrangidos por esta norma, pelos profissionais de saúde de departamento médico ou odontológico próprio ou conveniados às mesmas.

<tr><td class=titulo>13. ANOTAÇÕES NA CARTEIRA DE TRABALHO
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> está obrigada a promover, em quarenta e oito horas, às anotações nas Carteiras de Trabalho de seus AUXILIARES, ressalvados eventuais prazos mais amplos permitidos por lei.
<br><b>Parágrafo único</b> - É obrigatória a anotação na CTPS das mudanças provocadas por ascensão em plano de carreira ou alteração de titulação.

<tr><td class=titulo style="text-align:justify">14. MUDANÇA DE CARGO OU FUNÇÃO
<tr><td class=campo style="text-align:justify">O <b>AUXILIAR</b> não poderá ser transferido de um cargo ou função para outro, salvo com seu consentimento expresso e por escrito, sob pena de nulidade da referida transferência.

<tr><td class=titulo style="text-align:justify">15. ABONO DE FALTAS POR CASAMENTO OU LUTO
<tr><td class=campo style="text-align:justify">Não serão descontadas, no curso de nove dias corridos, as faltas do AUXILIAR, por motivo de gala ou luto, este em decorrência de falecimento de pai, mãe, filho(a), cônjuge, companheiro(a) e dependente juridicamente reconhecido.
<br><b>Parágrafo único</b> – Em caso de falecimento de irmão(ã), sogro(a) e neto(a) os abonos ficarão reduzidos a três dias.

<tr><td class=titulo style="text-align:justify">16. BOLSAS DE ESTUDO
<tr><td class=campo style="text-align:justify">Todo <b>AUXILIAR</b> que não esteja dentro do prazo do contrato de experiência tem direito a bolsas de estudo integrais, incluindo matrícula, no(s) estabelecimento(s) da <b>MANTENEDORA</b> localizado(s) no mesmo município onde leciona, conforme Instrução Normativa nº 15, de 06 de fevereiro de 2001, artigo 38, incisos I, II e II.
<br><b>Parágrafo 1º</b> - Somente terão direito a bolsas de estudo integrais, o(a) AUXILIAR, esposo(a) e companheiro(a), bem como seus filhos(as) e dependentes legais que estejam sob a guarda judicial, estes dois últimos desde que tenham 25 (vinte e cinco) anos ou menos na data de realização do exame vestibular ou do processo seletivo que define o ingresso no curso superior.
<br><b>Parágrafo 2º</b> - As bolsas de estudo integrais são válidas para cursos de graduação e seqüenciais existentes e administrados pela <b>MANTENEDORA</b> no(s) estabelecimento(s) de ensino superior localizado(s) no mesmo município para qual o <b>AUXILIAR</b> trabalha.
<br><b>Parágrafo terceiro</b> - A <b>MANTENEDORA</b> está obrigada, durante a vigência desta norma coletiva, a conceder duas bolsas de estudo integrais por AUXILIAR, no(s) estabelecimento(s) de ensino em que o mesmo trabalha, sendo que, nos cursos de graduação ou seqüenciais, não será possível que o bolsista conclua mais de um curso nesta condição.
<br><b>Parágrafo quarto</b> - A utilização do benefício previsto nesta cláusula é transitória e não habitual, por isso, não possui caráter remuneratório e nem se vincula, para nenhum efeito, ao salário ou remuneração percebida pelo AUXILIAR, nos termos do inciso XIX, do parágrafo 9º do artigo 214 do Decreto 3048, de 06 de maio de 1999 e do parágrafo 2º do artigo 458 da CLT, com a redação dada pela Lei 10243, de 19 de junho de 2001.
<br><b>Parágrafo 5º</b> - As bolsas de estudo integrais serão mantidas quando o <b>AUXILIAR</b> estiver licenciado para tratamento de saúde ou em gozo de licença mediante anuência da MANTENEDORA, de licenciamento para cumprimento de mandato sindical, nos termos do artigo 521, § único, da Consolidação das Leis do Trabalho, excetuados os casos de licença sem remuneração, para tratar de assuntos particulares.
<br><b>Parágrafo sexto</b> - No caso de falecimento do AUXILIAR, os dependentes que já se encontram estudando em estabelecimento de ensino superior da <b>MANTENEDORA</b> continuarão a gozar das bolsas de estudo integrais até o final do curso, ressalvado o disposto no parágrafo dez desta cláusula.
<br><b>Parágrafo sétimo</b> - No caso de dispensa sem justa causa durante o período letivo, ficam garantidas ao AUXILIAR, até o final do período letivo, as bolsas de estudo integrais já existentes.
<br><b>Parágrafo oitavo</b> - As bolsas de estudo integrais em cursos de pós-graduação ou de especialização existentes e administrados pela <b>MANTENEDORA</b> são válidas exclusivamente para o <b>AUXILIAR</b> em áreas correlatas àquelas em que o <b>AUXILIAR</b> exerce a função na <b>MANTENEDORA</b> e que visem à sua capacitação, respeitados os critérios de seleção exigidos para ingresso nos mesmos e obedecerão às seguintes condições:
<blockquote style="margin-top:0;margin-bottom:0">a) nos cursos stricto sensu ou de especialização que fixem um número máximo de alunos por turma, são limitadas em 30% (trinta por cento) do total de vagas oferecidas;
<br>b) nos cursos de pós-graduação lato sensu não haverá limites de vagas. Caso a estrutura do curso torne necessária a limitação do número de alunos será observado o disposto na alínea “a” deste parágrafo.</blockquote>
<br><b>Parágrafo nono</b> - As bolsas de estudos integrais concedidas nos termos do disposto no artigo 19 da lei nº 10.260 2001, poderão substituir, se for o caso, para as MANTENEDORAS de estabelecimentos de ensino superior sem fins lucrativos e beneficente de assistência social, o benefício tratado nesta cláusula.
<br><b>Parágrafo dez</b> - Os bolsistas que forem reprovados no período letivo perderão o direito à bolsa de estudo, voltando a gozar do benefício quando lograrem aprovação no referido período. As disciplinas cursadas em regime de dependência serão de total responsabilidade do bolsista, arcando o mesmo com o seu custo.
<br><b>Parágrafo onze</b> - Quando, a critério da MANTENEDORA, o AUXILIAR, em razão das funções exercidas na Instituição se vir na contingência de efetuar seus estudos, na área educacional indicada em outra instituição de ensino, a <b>MANTENEDORA</b> arcará com o valor integral das mensalidades do curso, incluindo matrícula durante a vigência do contrato de trabalho, respeitada a vigência coletiva de trabalho.
<br><b>Parágrafo doze</b> - Considera-se adquirido o direito daquele <b>AUXILIAR</b> que já esteja usufruindo bolsas de estudo em número superior ao definido nesta cláusula.
<br><b>Parágrafo treze</b> – O disposto nesta cláusula em seu caput e seus parágrafos, não se aplica ao <b>AUXILIAR</b> durante o contrato de experiência.

<tr><td class=titulo>17. IRREDUTIBILIDADE SALARIAL
<tr><td class=campo style="text-align:justify">É proibida a redução da remuneração mensal ou de carga horária do AUXILIAR, exceto quando ocorrer iniciativa expressa do mesmo. Em qualquer hipótese, é obrigatória a concordância formal e recíproca, firmada por escrito.
<br><b>Parágrafo Único</b> - Não havendo concordância recíproca, a parte que deu origem à redução prevista nesta cláusula arcará com a responsabilidade da rescisão contratual.

<tr><td class=titulo>18. UNIFORMES
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> deverá fornecer gratuitamente dois uniformes por ano, quando o seu uso for exigido.

<tr><td class=titulo>19. LICENÇA SEM REMUNERAÇÃO
<tr><td class=campo style="text-align:justify">O AUXILIAR, com mais de cinco anos ininterruptos de serviço no estabelecimento ensino superior da MANTENEDORA, terá direito a licenciar-se, sem direito à remuneração, por um período máximo de dois anos, não sendo este período de afastamento computado para contagem de tempo de serviço ou para qualquer outro efeito, inclusive legal.
<br><b>Parágrafo primeiro</b> - A licença ou sua prorrogação deverão ser comunicadas à <b>MANTENEDORA</b> com antecedência mínima de 90 (noventa) dias, devendo especificar as datas de início e término do afastamento. A licença só terá início a partir da data expressa no comunicado, mantendo-se, até aí, todas as vantagens contratuais. A intenção de retorno do <b>AUXILIAR</b> à atividade deverá ser comunicada à <b>MANTENEDORA</b> no mínimo 60 (sessenta) dias antes do término do afastamento.
<br><b>Parágrafo segundo</b> - O <b>AUXILIAR</b> que tenha ou exerça cargo de confiança deverá, junto com o comunicado de licença, solicitar seu desligamento do cargo a partir do início da licença.
<br><b>Parágrafo terceiro</b> - Considera-se demissionário o <b>AUXILIAR</b> que, ao término do afastamento, não retornar às atividades.

<tr><td class=titulo>20. LICENÇA À <b>AUXILIAR</b> ADOTANTE
<tr><td class=campo style="text-align:justify">Nos termos da Lei nº 10.421, de 15 de abril de 2.002, será garantida licença maternidade às <b>AUXILIARES</b> que vierem a adotar ou obtiverem guarda judicial de crianças.

<tr><td class=titulo>21. LICENÇA PATERNIDADE
<tr><td class=campo style="text-align:justify">A licença paternidade terá a duração de 5 dias.

<tr><td class=titulo>22. GARANTIA DE EMPREGO À GESTANTE
<tr><td class=campo style="text-align:justify">Fica garantido de emprego à <b>AUXILIAR</b> gestante desde o início da gravidez até sessenta dias após o término do afastamento legal. Em caso de dispensa, o aviso prévio começará a contar a partir do término do período de estabilidade.

<tr><td class=titulo>23. CRECHES
<tr><td class=campo style="text-align:justify">É obrigatória a instalação de local destinado à guarda de crianças de até seis anos, quando a unidade de ensino da <b>MANTENEDORA</b> mantiver contratadas, em jornada integral, pelo menos trinta funcionárias com idade superior a 16 anos. A manutenção da creche poderá ser substituída pelo pagamento do reembolso-creche, nos termos da legislação em vigor (CF, 7º, XXV, Artigo 389, parágrafo 1º da CLT e Portaria MTb nº 3296 de 03.09.86), ou ainda, a celebração de convênio com uma entidade reconhecidamente idônea.

<tr><td class=titulo>24. GARANTIAS AO <b>AUXILIAR</b> EM VIAS DE APOSENTADORIA
<tr><td class=campo style="text-align:justify">Fica assegurada ao <b>AUXILIAR</b> que, comprovadamente estiver a 24 meses ou menos da aposentadoria integral por tempo de serviço ou da aposentadoria por idade, a garantia de emprego durante o período que faltar até a aquisição do direito, exceto nos cargos de confiança ou de mandato com duração expressa de inicio e término.
<br><b>Parágrafo primeiro</b> - A garantia de emprego é devida ao <b>AUXILIAR</b> que esteja contratado pela <b>MANTENEDORA</b> há pelo menos três anos e que tenha comunicado à mesma a solicitação de sua contagem de tempo.
<br><b>Parágrafo segundo</b> - A comprovação à <b>MANTENEDORA</b> deverá ser feita mediante a apresentação de documento que ateste o tempo de serviço. Se o <b>AUXILIAR</b> depender de documentação para realização da contagem, terá um prazo de vinte e cinco dias, a contar da data da comunicação da dispensa. Comprovada a solicitação da documentação, os prazos serão prorrogados até que a mesma seja emitida. Este documento deverá ser emitido pela Previdência Social ou por funcionário credenciado junto ao órgão previdenciário.
<br><b>Parágrafo terceiro</b> - O contrato de trabalho do <b>AUXILIAR</b> só poderá ser rescindido por mútuo acordo homologado pela entidade sindical profissional, ou pedido de demissão, ou na ausência da entidade sindical profissional o contrato de trabalho poderá ser rescindido na Delegacia Regional do Trabalho.
<br><b>Parágrafo quarto</b> - Havendo acordo formal entre as partes, o <b>AUXILIAR</b> poderá exercer outra função compatível, durante o período em que estiver garantido pela estabilidade.
<br><b>Parágrafo quinto</b> - O aviso prévio, em caso de demissão sem justa causa, integra o período de estabilidade previsto nesta cláusula.

<tr><td class=titulo>25. MULTA POR ATRASO NA HOMOLOGAÇÃO DA RESCISÃO CONTRATUAL
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> deve homologar a rescisão contratual até o 20º dia após o pagamento das verbas rescisórias, conforme disposto no § 8º, do artigo 477, da CLT.
<br>O atraso na homologação obrigará a <b>MANTENEDORA</b> ao pagamento de multa, em favor do AUXILIAR, correspondente a um mês de sua remuneração. A partir do vigésimo dia de atraso, haverá ainda multa diária de 0,3% (três décimos percentuais) do salário mensal.
<br>A <b>MANTENEDORA</b> está desobrigada de pagar a multa quando o atraso vier a ocorrer, comprovadamente, por motivos alheios à sua vontade.
<br><b>Parágrafo único</b> – A entidade sindical profissional está obrigada a fornecer comprovante de comparecimento sempre que a <b>MANTENEDORA</b> se apresentar para homologação das rescisões contratuais e comprovar a convocação do AUXILIAR.

<tr><td class=titulo>26. DEMISSÃO POR JUSTA CAUSA
<tr><td class=campo style="text-align:justify">Quando houver demissão por justa causa, nos termos do art. 482, da CLT, a <b>MANTENEDORA</b> está obrigada a determinar na carta-aviso o motivo que deu origem à dispensa. Caso contrário, ficará descaracterizada a justa causa.

<tr><td class=titulo>27. READMISSÃO DO AUXILIAR
<tr><td class=campo style="text-align:justify">O <b>AUXILIAR</b> que for readmitido para a mesma função até doze meses após o seu desligamento ficará desobrigado de firmar contrato de experiência.

<tr><td class=titulo>28. INDENIZAÇÃO POR DISPENSA IMOTIVADA
<tr><td class=campo style="text-align:justify">O <b>AUXILIAR</b> demitido sem justa causa terá direito a indenizações, conforme as letras “a” e “b” a seguir colocadas, além do aviso prévio legal de trinta dias e das indenizações previstas nesta convenção, quando forem devidas, nas condições abaixo especificadas:
<blockquote style="margin-top:0;margin-bottom:0">a) 3 (três) dias para cada ano trabalhado na MANTENEDORA;
<br>b) aviso prévio adicional de (15) quinze dias, caso o <b>AUXILIAR</b> tenha, no mínimo, cinqüenta anos de idade e que, à data do desligamento, conte com pelo menos um ano de serviço na MANTENEDORA.</blockquote>
<br><b>Parágrafo primeiro</b> - Não estará obrigada ao pagamento da indenização, prevista na alínea “a”, a <b>MANTENEDORA</b> que tiver garantido ao <b>AUXILIAR</b> demitido, durante pelo menos um ano, pagamento mensal de adicional por tempo de serviço decorrente de plano de cargos e salários ou de anuênio, qüinqüênio ou equivalente, cujo valor corresponda a, no mínimo, 1% do valor do salário por ano trabalhado.
<br><b>Parágrafo segundo</b> - Não terá direito à indenização assegurada na alínea “b” do caput o <b>AUXILIAR</b> que, na data de admissão na MANTENEDORA, contar com mais de 50 (cinqüenta) anos de idade.
<br><b>Parágrafo terceiro</b> - Essas indenizações não contarão, para nenhum efeito como tempo de serviço.

<tr><td class=titulo>29. ATESTADOS DE AFASTAMENTO E SALÁRIOS
<tr><td class=campo style="text-align:justify">Sempre que solicitada, a <b>MANTENEDORA</b> deverá fornecer ao <b>AUXILIARES</b> atestado de afastamento e salário (AAS) previsto na legislação vigente.

<tr><td class=titulo>30. FÉRIAS
<tr><td class=campo style="text-align:justify">As férias dos <b>AUXILIARES</b> serão determinadas nos termos da legislação que rege a matéria, pela direção da MANTENEDORA, sendo admitida a compensação dos dias de férias concedidos antecipadamente, em período nunca inferior a dez dias e nem mais que duas vezes por ano.
<br><b>Parágrafo primeiro</b> – Fica assegurado aos <b>AUXILIARES</b> o pagamento, quando do início de suas férias, do salário correspondente às mesmas e do abono previsto no inciso XVII, artigo 7º , da Constituição Federal, no prazo previsto pelo artigo 145 da CLT, independentemente de solicitação pelos mesmos.
<br><b>Parágrafo segundo</b> – As férias, individuais ou coletivas, não poderão ter seu início coincidindo com domingos, feriados, dia de compensação do repouso semanal remunerado ou sábados, quando esses não forem dias normais de trabalho.

<tr><td class=titulo>31. DELEGADO REPRESENTANTE
<tr><td class=campo style="text-align:justify">Em cada unidade que tenha mais de 50 AUXILIARES, a <b>MANTENEDORA</b> assegurará eleição de um Delegado Representante, que terá garantia de emprego e salários a partir da inscrição de sua candidatura até seis meses após o término de sua gestão.
<br><b>Parágrafo primeiro</b> - O mandato do Delegado Representante será de um ano.
<br><b>Parágrafo segundo</b> - A eleição do Delegado Representante será realizada pela entidade sindical na unidade de ensino da MANTENEDORA, por voto direto e secreto. É exigido quorum de 50% (cinqüenta por cento) mais um dos <b>AUXILIARES</b> da unidade de ensino da <b>MANTENEDORA</b> onde a eleição ocorrer.
<br><b>Parágrafo terceiro</b> - A entidade sindical comunicará a eleição à MANTENEDORA, com antecedência mínima de sete dias corridos. Nenhum candidato poderá ser demitido a partir da data da comunicação até o término da apuração.
<br><b>Parágrafo quarto</b> - É condição necessária que os candidatos tenham, à data da eleição, pelo menos um ano de serviço na MANTENEDORA.

<tr><td class=titulo>32. QUADRO DE AVISOS
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> deverá colocar à disposição da entidade sindical da categoria profissional quadro de avisos, em local visível , para fixação de comunicados de interesse da categoria, sendo proibida a divulgação de matéria político-partidária ou ofensiva a quem quer que seja.

<tr><td class=titulo>33. ASSEMBLÉIAS SINDICAIS
<tr><td class=campo style="text-align:justify">Todo <b>AUXILIAR</b> terá direito a abono de faltas para o comparecimento às assembléias da categoria.
<br><b>Parágrafo primeiro</b> - Na vigência desta Convenção, os abonos estão limitados, a dois sábados e mais dois dias úteis, quando a assembléia não for realizada no município em que o <b>AUXILIAR</b> trabalhe para a MANTENEDORA. Caso a Assembléia ocorra fora do município em que o <b>AUXILIAR</b> trabalhe para MANTENEDORA, os abonos estão limitados, a dois sábados e dois períodos. As duas assembléias realizadas durante os dias úteis deverão ocorrer em períodos distintos.
<br><b>Parágrafo segundo</b> - A entidade sindical deverá informar à MANTENEDORA, por escrito, com antecedência mínima de quinze dias corridos. Na comunicação deverão constar a data e o horário da assembléia.
<br><b>Parágrafo terceiro</b> - Os dirigentes sindicais não estão sujeitos ao limite previsto no parágrafo primeiro desta cláusula. As ausências decorrentes do comparecimento às assembléias de suas entidades serão abonadas mediante comunicação formal à MANTENEDORA.
<br><b>Parágrafo quarto</b> - A <b>MANTENEDORA</b> poderá exigir dos <b>AUXILIARES</b> e dos dirigentes sindicais atestado emitido pela entidade sindical profissional ou pela FETEE, que comprove o seu comparecimento à assembléia.

<tr><td class=titulo>34. CONGRESSOS, SIMPÓSIOS E EQUIVALENTES
<tr><td class=campo style="text-align:justify">Os abonos de falta para comparecimento a congressos, simpósios e equivalentes serão concedidos mediante aceitação por parte da MANTENEDORA, que deverá formalizar por escrito a dispensa do AUXILIAR.
<br><b>Parágrafo único</b> - A participação do <b>AUXILIAR</b> nos eventos descritos no “caput” não caracterizará atividade extraordinária.

<tr><td class=titulo>35. CONGRESSO DA ENTIDADE SINDICAL PROFISSIONAL
<tr><td class=campo style="text-align:justify">Na vigência desta Convenção, a entidade sindical promoverá um evento de natureza política ou pedagógica (Congresso ou Jornada). A <b>MANTENEDORA</b> abonará as ausências de seus <b>AUXILIARES</b> que participarem do evento, nos seguintes limites:
<blockquote style="margin-top:0;margin-bottom:0">no estabelecimento de ensino superior que tenha até 49 AUXILIARES, será garantido, o abono a um AUXILIAR; 
<br>no estabelecimento de ensino superior que tenha entre 50 e 99 AUXILIARES, será garantido, o abono a dois AUXILIARES; 
<br>no estabelecimento de ensino superior que tenha mais de 100 AUXILIARES, será garantido, o abono a três AUXILIARES. </blockquote>
<br>Tais faltas, limitadas ao máximo de dois dias úteis além do sábado, serão abonadas mediante a apresentação de atestado de comparecimento fornecido pela entidade sindical ou pela FETEE. O <b>AUXILIAR</b> deverá repor as horas que, porventura, sejam necessárias para complementação da sua jornada de trabalho.

<tr><td class=titulo>36. RELAÇÃO NOMINAL
<tr><td class=campo style="text-align:justify">Obriga-se a <b>MANTENEDORA</b> a encaminhar para entidade representativa da categoria profissional, conforme Precedentes Normativos nºs 41 e 111, do Tribunal Superior do Trabalho, no prazo máximo de trinta dias contados da data do recolhimento da Contribuição Sindical, a relação nominal dos <b>AUXILIARES</b> que integram seu quadro de funcionários acompanhada do valor do salário mensal e das guias das contribuições sindical e assistencial.

<tr><td class=titulo>37. FORO CONCILIATÓRIO PARA SOLUÇÃO DE CONFLITOS COLETIVOS
<tr><td class=campo style="text-align:justify">Fica mantida a existência do Foro Conciliatório para Solução de Conflitos Coletivos, que tem como objetivo procurar resolver:
<blockquote style="margin-top:0;margin-bottom:0">I - divergências trabalhistas;
<br>II - incapacidade econômico-financeira da MANTENEDORA, no cumprimento de reajuste salarial e/ou de cláusulas previstas na presente convenção coletiva;
<br>III – alteração no prazo de pagamento de salários.</blockquote>
    <b>Parágrafo primeiro</b> - Havendo dificuldade no cumprimento da cláusula de reajuste salarial ou diminuição nos percentuais de reajustes salariais estipulados nesta convenção coletiva ou definição de outro critério de reajuste salarial proposto pela MANTENEDORA, a solicitação da realização do Foro deverá ser formalizada por escrito e instruída com a documentação pertinente ao pedido.
<br><b>Parágrafo segundo</b> - Para efeito do que estabelece os incisos I, II e III deste artigo, a MANTENEDORA, ao solicitar o FORO, deve encaminhar os motivos do pedido de liberação do cumprimento da cláusula em questão, acompanhada da competente documentação comprobatória, para análise e decisão.
<br><b>Parágrafo terceiro</b> - O Foro será composto paritariamente, por três representantes do SEMESP, da FETEE e da entidade representativa da categoria profissional. As reuniões deverão contar, também, com as partes em conflito que, se assim o desejarem, poderão delegar representantes para substituí-las e/ou serem assistidas por advogados, com poderes específicos para adotarem, em nome da Instituição, as decisões julgadas convenientes e necessárias.
<br><b>Parágrafo quarto</b> - O SEMESP, a FETEE e a entidade representativa da categoria profissional deverão indicar os seus representantes no Foro num prazo de trinta dias a contar da assinatura desta Convenção.
<br><b>Parágrafo quinto</b> - Cada sessão do Foro será realizada no prazo máximo de quinze dias a contar da solicitação formal e obrigatória de qualquer uma das entidades que o compõem. A data, o local e o horário serão decididos pelas entidades sindicais envolvidas. O não comparecimento de qualquer uma das partes acarretará no encerramento imediato das negociações, bem como na aplicação na multa estabelecida no Parágrafo nono desta cláusula.
<br><b>Parágrafo sexto</b> - Nenhuma das partes envolvidas ingressará com ação na Justiça do Trabalho durante as negociações de entendimento.
<br><b>Parágrafo sétimo</b> - Na ausência de solução do conflito ou na hipótese de não comparecimento de qualquer uma das partes, a comissão responsável pelo Foro fornecerá certidão atestando o encerramento da negociação.
<br><b>Parágrafo oitavo</b> - Na hipótese de sucesso das negociações, a critério do Foro, a <b>MANTENEDORA</b> ficará desobrigada de arcar com a multa prevista no item 9 º (nono) desta cláusula.
<br><b>Parágrafo nono</b> - As decisões do Foro terão eficácia legal entre as partes acordantes. O descumprimento das decisões assumidas gerará multa a ser estabelecida no Foro, independentemente daquelas já estabelecidas nesta Convenção.
<br><b>Parágrafo dez</b> - A entidade sindical ou a <b>MANTENEDORA</b> que deixar de comparecer ao FORO, uma vez convocada, pagará uma multa de R$ 1.000,00 (hum mil reais), que reverterá em favor da parte presente.

<tr><td class=titulo>38. COMISSÃO PERMANENTE DE NEGOCIAÇÃO
<tr><td class=campo style="text-align:justify">Fica mantida a Comissão Permanente de Negociação constituída de forma paritária, por três (3) representantes das entidades sindicais profissionais e econômica, com o objetivo de:
<blockquote style="margin-top:0;margin-bottom:0">a) fiscalizar o cumprimento das cláusulas vigentes;
<br>b) elucidar eventuais divergências de interpretação das cláusulas desta Convenção;
<br>c) discutir questões não-contempladas na norma coletiva;
<br>d) deliberar, no prazo máximo de trinta dias a contar da data da solicitação protocolizada no SEMESP, sobre a isenção prevista na cláusula referente às indenizações por dispensa imotivada constante da presente Convenção e sobre modificação de pagamento da assistência médico-hospitalar, conforme os parágrafos 1º e 3º da cláusula relativa à matéria, constante desta norma coletiva;
<br>e) criar subsídios para a Comissão de Tratativas Salariais 2005/2006, através da elaboração de documentos para a definição das funções/atividades e o regime de trabalho dos AUXILIARES.
<br>f) criar critérios para a regionalização das negociações salariais referentes a 2004, bem como definir critérios diferenciados para elaboração do instrumento normativo destinado às entidades mantenedoras de Universidades, Centros Universitários, Faculdades, Institutos Superiores de Educação e Centros de Educação Tecnológicas.</blockquote>
    <b>Parágrafo primeiro</b> – As entidades sindicais componentes da Comissão Permanente de Negociação indicarão seus representantes, no prazo máximo de trinta dias corridos, a contar da assinatura da presente Convenção.
<br><b>Parágrafo segundo</b> – A Comissão Permanente de Negociação deverá reunir-se mensalmente, em calendário elaborado de comum acordo entre as partes, alternadamente nas sedes das entidades sindicais que a compõem. Nos casos dispostos na letra “d” do caput, deverá haver convocação específica pela entidade sindical patronal.
<br><b>Parágrafo terceiro</b> - O não comparecimento da entidade sindical, profissional ou econômica, nas reuniões previstas no Parágrafo segundo da presente cláusula, implicará na multa de R$ 2.000,00 (dois mil reais) por reunião, a qual reverterá em benefício da entidade presente à mesma.

<tr><td class=titulo>39. ACORDOS INTERNOS
<tr><td class=campo style="text-align:justify">Ficam assegurados os direitos mais favoráveis decorrentes de acordos internos ou de acordos coletivos de trabalho celebrados entre a <b>MANTENEDORA</b> e a entidade sindical profissional.

<tr><td class=titulo>40. ASSISTÊNCIA MÉDICO-HOSPITALAR
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> está obrigada a assegurar, às suas expensas, assistência médico-hospitalar a todos os seus AUXILIARES, sendo-lhe facultada a escolha por plano de saúde, seguro-saúde ou convênios com empresas prestadoras de serviços médico-hospitalares. Poderá, ainda, prestar a referida assistência diretamente em se tratando de instituições que disponham de serviços de saúde e hospitais próprios ou conveniados. Qualquer que seja a opção feita, a assistência médico-hospitalar deve assegurar as condições e os requisitos mínimos que seguem relacionados:
<blockquote style="margin-top:0;margin-bottom:0">1. Abrangência – A assistência médico-hospitalar deve ser realizada no município onde funciona o estabelecimento de ensino superior ou onde vive o AUXILIAR, a critério da MANTENEDORA. Em casos de emergência, deverá haver garantia de atendimento integral em qualquer localidade do Estado de São Paulo ou fixação, em contrato, de formas de reembolso.
<br>2. Coberturas mínimas:
<blockquote style="margin-top:0;margin-bottom:0">2.1 Quarto para quatro pacientes, no máximo.
<br>2.2 Consultas.
<br>2.3 Prazo de internação de 365 dias por ano (comum e UTI/CTI)
<br>2.4 Parto, independentemente do estado gravídico.
<br>2.5 Moléstias infecto-contagiosas que exijam internação.
<br>2.6 Exames laboratoriais, ambulatoriais e hospitalares.</blockquote>
3. Carência – Não haverá carência na prestação dos serviços médicos e laboratoriais.
<br>4. <b>AUXILIAR</b> ingressante – Não haverá carência para o <b>AUXILIAR</b> ingressante, independentemente do mês em que for contratado.
<br>5. Pagamento</blockquote>
<blockquote style="margin-top:0;margin-bottom:0">A assistência médico-hospitalar será garantida nos termos desta Convenção, cabendo ao AUXILIAR, para usufruir dos benefícios da Lei nº 9656/98, o pagamento de 10% das mensalidades da referida assistência, com teto limite de R$ 8,00 (oito reais) por mês, respeitado o estabelecido no parágrafo 1º desta cláusula.</blockquote>
    <b>Parágrafo primeiro</b> – Caso a assistência médico-hospitalar vigente na Instituição venha a sofrer reajuste em virtude de possíveis modificações estabelecidas em legislação que abranja o segmento – Lei 9.656, de 03 de junho de 1998 e MP 2.097-39, de 26 de abril de 2001 - ou que vierem a ser estabelecidas em lei, ou por mudança de empresa prestadora de serviço, a pedido do corpo técnico-administrativo da Instituição ou por quebra de contrato, unilateralmente, por parte da atual empresa prestadora de serviço, a <b>MANTENEDORA</b> continuará a contribuir com o valor mensal vigente até a data da modificação, devendo o AUXILIARarcar com o valor excedente, que será descontado em folha e consignado no comprovante de pagamento, nos termos do art. 462, da CLT.
<br><b>Parágrafo segundo</b> - Caso ocorra mudança de empresa prestadora de serviço, por decisão unilateral da MANTENEDORA, com conseqüente reajuste no valor vigente, o <b>AUXILIAR</b> estará isento do pagamento do valor excedente, cabendo à <b>MANTENEDORA</b> prover integralmente a assistência médico-hospitalar, sem nenhum ônus para o AUXILIAR.
<br><b>Parágrafo terceiro</b> – Para efeito do disposto no Parágrafo primeiro desta cláusula, caberá à <b>MANTENEDORA</b> remeter a documentação comprobatória à Comissão Permanente de Negociação, nos termos do artigo 47, da presente norma, para a devida homologação.
<br><b>Parágrafo quarto</b> – Fica obrigado o <b>AUXILIAR</b> a optar pela prestação de assistência médico-hospitalar em uma única Instituição de ensino, quando mantiver mais de um vínculo empregatício como AUXILIARno mesmo município ou municípios conurbanos. É necessário que o <b>AUXILIAR</b> se manifeste por escrito, com antecedência mínima de vinte dias, para que a <b>MANTENEDORA</b> possa proceder à suspensão dos serviços.
<br><b>Parágrafo quinto</b> – Mediante pagamento complementar e adesão facultativa, conforme o plano de atendimento médico-hospitalar e devidamente documentado, o <b>AUXILIAR</b> poderá optar pela ampliação dos serviços de saúde garantidos nesta Convenção Coletiva ou estendê-los a seus dependentes.

<tr><td class=titulo>41. SALÁRIO DO <b>AUXILIAR</b> ADMITIDO PARA SUBSTITUIÇÃO
<tr><td class=campo style="text-align:justify">Ao <b>AUXILIAR</b> admitido em substituição a outro desligado, qualquer que tenha sido o motivo do seu desligamento, será garantido, sempre, salário inicial igual ao menor salário na função existente no estabelecimento, curso, grau ou nível de ensino, respeitado o Plano de Cargos e Salários da MANTENEDORA, sem serem consideradas eventuais vantagens pessoais.

<tr><td class=titulo>42. MENOR SALÁRIO DA CATEGORIA
<tr><td class=campo style="text-align:justify">Fica assegurado, a partir de 1º (primeiro) de maio de 2005, nos termos do inciso V, artigo 7º, da Constituição Federal, um menor salário da categoria equivalente a R$ 490,92 (quatrocentos e noventa reais e noventa e dois centavos) por jornada integral de trabalho (44 horas semanais).
<br><b>Parágrafo único</b> – Para o ano de 2006, o menor salário da categoria consignado no caput, será reajustado na conformidade do estabelecido na cláusula terceira da presente norma coletiva.

<tr><td class=titulo>43. ABONO DE PONTO AO ESTUDANTE
<tr><td class=campo style="text-align:justify">Fica assegurado o abono de faltas ao <b>AUXILIAR</b> estudante para prestação de exames escolares, condicionado à prévia comunicação à <b>MANTENEDORA</b> e comprovação posterior.

<tr><td class=titulo>44. PRORROGAÇÃO DA JORNADA DO ESTUDANTE
<tr><td class=campo style="text-align:justify">Fica permitida a prorrogação da jornada de trabalho ao "AUXILIAR" estudante, ressalvadas as hipóteses de conflito com horário de freqüência às aulas.

<tr><td class=titulo>45. ESTABILIDADE PROVISÓRIA DO ALISTANDO
<tr><td class=campo style="text-align:justify">É assegurada aos <b>AUXILIARES</b> em idade de prestação do serviço militar estabilidade provisória, desde o alistamento até sessenta dias após a baixa.

<tr><td class=titulo>46. <b>AUXILIAR</b> AFASTADO POR DOENÇA
<tr><td class=campo style="text-align:justify">Ao <b>AUXILIAR</b> afastado do serviço por doença devidamente atestada pela Previdência Social ou por médico ou dentista credenciado pela MANTENEDORA, será garantido o emprego ou o salário, a partir da alta, por igual período ao do afastamento, limitado a 60 (sessenta) dias além do aviso prévio.

<tr><td class=titulo>47. REFEITÓRIOS
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> que contar com mais de 300 (trezentos) AUXILIARESno mesmo estabelecimento de ensino superior por ela mantido e não conceder vale-refeição, obriga-se a manter refeitório.
<br><b>Parágrafo único</b> – No estabelecimento de ensino superior da <b>MANTENEDORA</b> em que trabalhem menos de 300 (trezentos) <b>AUXILIARES</b> será obrigatório assegurar-lhes condições de conforto e higiene por ocasião das refeições.

<tr><td class=titulo style="text-align:justify">48. CESTA BÁSICA
<tr><td class=campo style="text-align:justify">Fica assegurada aos <b>AUXILIARES</b> que percebam, até 5 (cinco) salários mínimos por mês, em jornada integral de 44 (quarenta e quatro) horas semanais, a concessão de uma cesta básica mensal de 26 kg, composta, no mínimo, dos seguintes produtos não perecíveis:
<div align="center"><table width=350>
<tr><td class=campo>Arroz            </td><td class=campo>Óleo                </td><td class=campo>Macarrão </td></tr>
<tr><td class=campo>Feijão           </td><td class=campo>Café                </td><td class=campo>Sal </td></tr>
<tr><td class=campo>Farinha de Trigo </td><td class=campo>Farinha de Mandioca </td><td class=campo>Farinha de Milho </td></tr>
<tr><td class=campo>Açúcar           </td><td class=campo>Biscoito            </td><td class=campo>Purê de Tomate </td></tr>
<tr><td class=campo>Tempero          </td><td class=campo>Achocolatado        </td><td class=campo>Leite em Pó </td></tr>
<tr><td class=campo>Fubá             </td><td class=campo>Sardinha em Lata    </td><td class=campo>Sopão </td></tr>
</table></div>
    <b>Parágrafo primeiro</b> - As MANTENEDORAS que já concedem vale-refeição, conforme o determinado pelo PAT, estão desobrigadas do fornecimento de cesta básica.
<br><b>Parágrafo segundo</b> - Fica assegurada a concessão de cesta básica durante as férias, licença maternidade e licença doença, bem como será garantido ao <b>AUXILIAR</b> demitido sem justa causa, na vigência da presente Convenção, a cesta básica referente ao período de aviso prévio, ainda que indenizado.

<tr><td class=titulo>49. COMPENSAÇÃO SEMANAL DA JORNADA DE TRABALHO
<tr><td class=campo style="text-align:justify">Fica permitida a compensação semanal da jornada de trabalho, nos termos da Legislação que rege a matéria e obedecido o seguinte critério:
a) mediante ciência, através do calendário anual a ser publicado pela MANTENEDORA, os <b>AUXILIARES</b> serão dispensados do cumprimento de sua jornada de trabalho em dias ali previstos, compensando-se as horas não trabalhadas com horas de trabalho complementares. 

<tr><td class=titulo>50. BANCO DE HORAS
<tr><td class=campo style="text-align:justify">Nos termos da Lei nº 9.601, de 21 de janeiro de 1998, fica celebrado o Banco de Horas entre os <b>AUXILIARES</b> e as MANTENEDORAS, conforme documento anexo a presente CCT.
<br><b>Parágrafo primeiro</b> - As MANTENEDORAS que desejarem implantar o Banco de Horas, conforme o disposto no caput, deverão comunicar à entidade representativa da categoria profissional a implantação do mesmo, sob pena de não o fazendo não ter validade a aplicabilidade do Banco de Horas.
<br><b>Parágrafo segundo</b> - Caso a <b>MANTENEDORA</b> queira fazer alterações no Banco de Horas devido as suas peculiaridades, os critérios, detalhes, prazos e datas de implantação serão objeto de Acordo Coletivo de Trabalho específico, firmado entre a <b>MANTENEDORA</b> e seus AUXILIARES, com a participação da entidade sindical representativa da categoria profissional, na forma da legislação em vigor.

<tr><td class=titulo>51. AUTORIZAÇÃO PARA DESCONTO EM FOLHA DE PAGAMENTO
<tr><td class=campo style="text-align:justify">O desconto do <b>AUXILIAR</b> em folha de pagamento somente poderá ser realizado, mediante sua autorização, nos termos dos artigos 462 e 545, da CLT, quando os valores forem destinados ao custeio de prêmios de seguro, planos de saúde, mensalidades associativas ou outras que constem da sua expressa autorização, desde que não haja previsão expressa de desconto na presente norma coletiva.
<br><b>Parágrafo único</b> – Encontra-se na entidade sindical profissional, à disposição da MANTENEDORA, cópia de autorização do <b>AUXILIAR</b> para o desconto da mensalidade associativa.

<tr><td class=titulo>52. ESTABILIDADE PARA PORTADORES DE DOENÇAS GRAVES
<tr><td class=campo style="text-align:justify">Aos <b>AUXILIARES</b> acometidos por doenças graves ou incuráveis e aos <b>AUXILIARES</b> portadores do vírus HIV que vierem a apresentar qualquer tipo de infecção ou doença oportunista, resultante da patologia de base, não sendo julgados aptos para o trabalho por exame médico circunstanciado, fica assegurada estabilidade até encaminhamento de pedido ao órgão previdenciário para gozar do benefício saúde ou até a eventual concessão de aposentadoria por invalidez.
<br><b>Parágrafo único</b> – São consideradas doenças graves ou incuráveis, a tuberculose ativa, alienação mental, esclerose múltipla, neoplasia maligna, cegueira definitiva, hanseníase, cardiopatia grave, doença de Parkinson, paralisia irreversível e incapacitante, espondiloastrose anquilosante, neofropatia grave, estados do Mal de Paget (osteíte deformante) e contaminação grave por radiação.

<tr><td class=titulo>53. NÚCLEO INTERSINDICAL DE CONCILIAÇÃO TRABALHISTA
<tr><td class=campo style="text-align:justify">Poderá ser criado, nas localidades onde já não esteja instalado, o Núcleo Intersindical de Conciliação Trabalhista que funcionará no sentido de buscar a composição de conflitos no âmbito das relações entre as partes representadas pelas entidades signatárias desta Convenção, nos termos previstos pelo artigo 625-C da Consolidação das Leis do Trabalho, com a redação dada pela Lei 9.958, de 12 de janeiro de 2000.

<tr><td class=titulo>54. GARANTIAS AO <b>AUXILIAR</b> COM SEQUELAS E READAPTAÇÃO
<tr><td class=campo style="text-align:justify">Será garantida ao <b>AUXILIAR</b> acidentado no trabalho ou acometido por doença profissional, a permanência na <b>MANTENEDORA</b> em função compatível com seu estado físico, sem prejuízo da remuneração antes percebida, desde que após o acidente ou comprovação da aquisição de doença profissional apresente, cumulativamente, redução da capacidade laboral, atestada por órgão oficial e que se tenha tornado incapaz de exercer a função que anteriormente desempenhava, obrigado, porém, o <b>AUXILIAR</b> nessa situação a participar dos processos de readaptação e reabilitação profissionais.
<br><b>Parágrafo único</b> – O período de estabilidade do <b>AUXILIAR</b> que se encontra participando dos processos de readaptação e reabilitação profissionais será o previsto em lei.

<tr><td class=titulo>55- COMPETÊNCIA DAS ENTIDADES SINDICAIS SIGNATÁRIAS
<tr><td class=campo style="text-align:justify">Fica estabelecida a legalidade das entidades sindicais signatárias para promover, perante a Justiça do Trabalho e o Foro em Geral, ações plúrimas em nome dos <b>AUXILIARES</b> em nome próprio, ou ainda, como parte interessada, em caso de descumprimento de qualquer cláusula avençada ou determinada nesta norma coletiva.

<tr><td class=titulo>56- PRIMEIROS SOCORROS
<tr><td class=campo style="text-align:justify">A <b>MANTENEDORA</b> obriga-se a manter materiais de primeiros socorros nos locais de trabalho e providenciar, por sua conta, a remoção do <b>AUXILIAR</b> acidentado/doente para o atendimento médico-hospitalar.

<tr><td class=titulo>57 – FLEXIBILIZAÇÃO DA JORNADA DE TRABALHO
<tr><td class=campo style="text-align:justify">Poderá ser flexibilizada a carga horária entre jornadas do AUXILIAR, quando no exercício concomitante de função docente e atividade administrativa, não havendo assim pagamento de salários nos intervalos, quando o <b>AUXILIAR</b> não tenha trabalhado nos mesmos.

<tr><td class=titulo>58. MULTA POR DESCUMPRIMENTO DA CONVENÇÃO
<tr><td class=campo style="text-align:justify">O descumprimento desta Convenção obrigará a <b>MANTENEDORA</b> ao pagamento de multa correspondente a 5% (cinco por cento) do salário do AUXILIAR, acrescida de juros e correção monetária, para cada <b>AUXILIAR</b> prejudicado.
<br><b>Parágrafo único</b> - A <b>MANTENEDORA</b> está desobrigada de arcar com o valor previsto nesta cláusula, caso o artigo da Convenção já estabeleça uma multa pelo não cumprimento da mesma.

<tr><td class=titulo>59. Contribuição assistencial profissional – SAAE/ABC
<tr><td class=campo style="text-align:justify">Considerando o disposto no artigo 8º, inciso I, da Constituição Federal “que veda ao Poder Público a interferência e a intervenção na organização sindical”; 
<br>Considerando o disposto no artigo 7º, inciso XXVI, da Carta Maior “reconhece as convenções e os acordos coletivos de trabalho”; 
<br>Considerando o disposto no artigo 613 e parágrafos da Consolidação das Leis do Trabalho e incisos que estabelece “terem as convenções e os acordos coletivos de trabalho efeito “erga omnes”; 
<br>Considerando o disposto no artigo 614 e parágrafos do texto consolidado que “determina que as convenções e os acordos coletivos de trabalho, após três dias da entrega dos mesmos no órgão competente do Ministério do Trabalho e Emprego, entram em vigor, fazendo lei entre as partes”; 
<br>Considerando o disposto no artigo 8º, inciso III, da Lei Magna, que estabelece “ao sindicato cabe a defesa dos direitos e interesses coletivos e individuais da categoria, inclusive em questões judiciais ou administrativas”; 
<br>Considerando o disposto no artigo 8º, da Convenção 95, da Organização Internacional do Trabalho (OIT), da qual o Brasil é signatário e, portanto, obrigado, que estabelece “descontos em salários não serão autorizados, senão sob condições e limites prescritos pela legislação nacional ou fixados por convenções coletivas de trabalho ou sentença arbitral”; 
<br>Considerando o disposto no Verbete nº 324, do Comitê de Liberdade Sindical, da Organização Internacional do Trabalho, do qual o Brasil é signatário e, portanto, obrigado, que estabelece “obrigação do pagamento da quota de solidariedade dos não filiados em relação aos filiados, como condição para que tenham as vantagens estabelecidas nos Instrumentos Normativos”; 
<br>Considerando que o Supremo Tribunal Federal, em 7/11/2000, no Processo RE 189960-SP, decidiu, conforme Certidão de Julgamento que “A Turma entendeu que é legítima a cobrança de contribuição assistencial imposta aos empregados indistintamente em favor do sindicato, prevista em convenção coletiva de trabalho, estando os não sindicalizados compelidos a satisfazer a mencionada contribuição”; 
<br>Considerando que o mesmo Supremo Tribunal Federal, no julgamento do Agravo Regimental interposto no R.E. nr 337718, em 1º/8/2002, sendo relator o Excelentíssimo Senhor Ministro Nelson Jobim, prolatou a seguinte EMENTA – CONTRIBUIÇÃO COLETIVA: “A contribuição prevista em convenção coletiva, fruto do disposto no artigo 513, alínea “e”, da Constituição Federal é devida por todos os integrantes da categoria profissional, não se confundindo com aquela versada na primeira parte do inciso IV, do artigo 8º, da Carta da República. (r.e. 189960, Marco Aurélio, DJ 10/08/2001). “Estive presente ao julgamento do referido recurso. “Acompanhei Marco Aurélio”. Coerente com a posição tomada, dou provimento ao regimental para conhecer e prover integralmente o RE do Sindicato dos Metalúrgicos do ABC e outros”. Publique-se. Brasília, 1. de agosto de 2002. Ministro Nelson Jobim, Relator. 
<br><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 7% (sete por cento), em duas parcelas de 3,5% do salário mensal bruto de cada “AUXILIAR”, para desconto nos meses de junho e outubro e recolhimento até o dia 15 do respectivo mês subseqüente, observado o teto-limite de R$ 200,00 por vez, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.
<br><b>§ 2º</b> - O recolhimento será feito obrigatoriamente pela própria MANTENEDORA, até o dia 15 do mês subsequente ao desconto, em guias próprias enviadas pela entidade sindical profissional, acompanhadas das competentes relações nominais e valores devidos. Essas importâncias destinam-se à manutenção e ampliação dos serviços assistenciais da entidade sindical profissional, bem como a permitir a participação da mesma nas negociações com os sindicatos patronais.
<br><b>§ 3º</b> - Quando a <b>MANTENEDORA</b> deixar de efetuar o desconto e o recolhimento das contribuições estabelecidas nesta cláusula, decorrentes da decisão da assembléia geral da categoria profissional, incorrerá na obrigatoriedade do pagamento de multa, cujo valor corresponderá a 5% (cinco por cento) do total da importância a ser recolhida para a entidade sindical representativa da categoria profissional, acrescida da parcela correspondente à variação da TR ou de outro índice que vier a substituí-la, a partir do dia seguinte ao do vencimento, cabendo à <b>MANTENEDORA</b> a integral responsabilidade pela multa e demais cominações, não podendo as mesmas, de forma alguma, incidir sobre os salários dos AUXILIARES.
<br><b>§ 4º</b> - O desconto e o recolhimento da contribuição assistencial, bem como os respectivos valores, foram decididos, com base nos textos legais acima mencionados, em assembléia geral especificamente convocada e amplamente divulgada através de editais publicados em 34 (trinta e quatro) jornais de grande circulação estadual e regional e devidamente realizada, nos termos do artigo 513, “e”, da Consolidação das Leis do Trabalho, que estabelece, como prerrogativa das entidades sindicais “impor contribuições a todos aqueles que participam das categorias econômicas ou profissionais ou das profissões liberais representadas”.

<tr><td class=titulo>59. Contribuição assistencial profissional – Araçatuba e Região
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 10% (dez por cento), em 10 parcelas de 1% do salário mensal bruto de cada “AUXILIAR”, para desconto a partir do mês de junho e assim sucessivamente até completar as dez parcelas e recolhimento até o dia 15 do respectivo mês subseqüente, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. Contribuição assistencial profissional – BAURU
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 6% (seis por cento), em três parcelas de 2% do salário mensal bruto de cada “AUXILIAR”, para desconto nos meses de junho, julho e agosto e recolhimento até o dia 15 do respectivo mês subseqüente, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. Contribuição assistencial profissional – dracena e região
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 5% (cinco por cento), em duas parcelas de 2,5% do salário mensal bruto de cada “AUXILIAR”, para desconto até o dia 30 de agosto e 30 de novembro, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL – FERNANDÓPOLIS/ JALES
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 10%, em duas parcelas de 5% do salário mensal bruto de cada “AUXILIAR”, para desconto nos meses de junho e outubro, para recolhimento até o dia 15 do respectivo mês subseqüente,conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL – LINS
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 7%, em duas parcelas de 3,5% do salário mensal bruto de cada AUXILIAR, para desconto nos meses de junho (recolhida até 08/07) e outubro (recolhida até 11/11), observado o teto-limite de R$ 150,00 por vez, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL – MARÍLIA
<tr><td class=campo style="text-align:justify">Não tem contribuição Assistencial

<tr><td class=titulo>59. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL – MOGI DAS CRUZES
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 7% para desconto em duas parcelas de 3,5% do salário mensal bruto de cada AUXILIAR, nos meses de junho e julho, para recolhimento até o dia 10 de cada mês subseqüente, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL – PIRACICABA
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 6%, em três parcelas de 2% do salário mensal bruto de cada AUXILIAR, a serem descontadas nos meses de junho, setembo e novembro, para recolhimento até o dia 15 do respectivo mês subseqüente, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL – PRESIDENTE PRUDENTE
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 7%, em duas parcelas de 3,5% do salário mensal bruto de cada “AUXILIAR”, para desconto nos meses de junho e novembro, para recolhimento até o dia 10 do respectivo mês subseqüente, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL – RIBEIRÃO PRETO
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 10%, em duas parcelas de 5% do salário mensal bruto de cada AUXILIAR, a serem descontadas nos meses de junho e setembro, para recolhimento até o dia 15 do respectivo mês subseqüente, observado o teto-limite de R$ 50,00 por vez, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL – SANTOS
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 5% (sete por cento) do salário mensal bruto de cada AUXILIAR, para desconto no mês de junho e recolhimento até o dia 15 do respectivo mês subseqüente, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=titulo>59. CONTRIBUIÇÃO ASSISTENCIAL PROFISSIONAL – SÃO JOSÉ DO RIO PRETO/ SOROCABA
<tr><td class=campo style="text-align:justify"><b>§ 1º</b> - Obrigam-se as MANTENEDORAS a promoverem, no exercício de 2005, na folha de pagamento dos seus “AUXILIARES” sindicalizados e/ou filiados ou não, para recolhimento em favor da entidade sindical signatária, legalmente representativa da categoria na base territorial conferida à mesma pela respectiva Carta Sindical ou Registro definitivo no Cadastro Nacional das Entidades Sindicais (CNES) do Ministério do Trabalho e Emprego, o desconto da importância correspondente a 10% (cinco por cento), em duas parcelas de 5% do salário mensal bruto de cada AUXILIAR, para desconto nos meses de junho e novembro, para recolhimento até o dia 15 do respectivo mês subseqüente, observado o teto-limite de R$ 50,00 por vez, conforme estabelecido na assembléia geral da categoria, a título de contribuição assistencial.

<tr><td class=campo style="text-align:justify">Por estarem justos e acertados, assinam a presente Convenção Coletiva de Trabalho de 2005, a qual será depositada, para fins de arquivo, na Delegacia Regional do Trabalho e Emprego no Estado de São Paulo, nos termos do artigo 614, da Consolidação das Leis do Trabalho, de modo a surtir, de imediato, os seus efeitos legais.

<tr><td class=campo style="text-align:justify">São Paulo, junho de 2005.
<br>
<br>Hermes Ferreira Figueiredo
<br>Presidente do SEMESP
<br>CPF/MF nº 04.946.158-34
<br>
<br>Geraldo Mugayar
<br>Federação dos Trabalhadores em Estabelecimentos de Ensino do Estado de São Paulo - FETEE
<br>CPF/MF nº 023.779.778-04
<br>
<br>José Roberto Marques de Castro
<br>Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Marília
<br>
<br>Ronaldi Torelli
<br>Sindicato dos Professores e Trabalhadores em Educação de Dracena e Região
<br>
<br>Ayrton Onofre da Silva
<br>Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Lins
<br>
<br>Ademir Rodrigues
<br>Sindicato dos Trabalhadores em Estabelecimentos de Ensino de Presidente Prudente
<br>
<br>Rita Theresinha de Miranda Furquim
<br>Sindicato dos Professores e <b>AUXILIARES</b> de Administração Escolar de Ribeirão Preto
<br>
<br>Luiz Carlos Custódio
<br>Sindicato dos Professores e <b>AUXILIARES</b> Administrativos de Araçatuba e Região
<br>
<br>Celso Soares Nogueira
<br>Sindicato dos <b>AUXILIARES</b> de Administração Escolar do ABC
<br>
<br>Fátima Aparecida Marins Silva
<br>Sindicato dos <b>AUXILIARES</b> de Administração Escolar de Bauru
<br>
<br>José Cláudio Chaves
<br>Sindicato dos <b>AUXILIARES</b> de Administração Escolar de Mogi das Cruzes
<br>
<br>João Manoel dos Santos
<br>Sindicato dos <b>AUXILIARES</b> de Administração Escolar de Piracicaba
<br>
<br>Márcio Campos
<br>Sindicato dos <b>AUXILIARES</b> de Administração Escolar de Santos
<br>
<br>Cláudio Figueroba Raimundo
<br>Sindicato dos <b>AUXILIARES</b> de Administração Escolar de Sorocaba
<br>
<br>Valdecir Zampolla Caetano
<br>Sindicato dos <b>AUXILIARES</b> de Administração Escolar de São José do Rio Preto
<br>CPF/MF nº 025.666.518-41 
<br>
</table>

<DIV style="page-break-after:always"></DIV>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td class=titulo align="center">ANEXO 01
<tr><td class=titulo align="center">ACORDO COLETIVO DE TRABALHO PARA A INSTITUIÇÃO DE BANCO DE HORAS. 
<tr><td class=campo style="text-align:justify"><b>Cláusula Primeira</b> – Fica estabelecido entre as MANTENEDORAS, neste ato representadas pelo SEMESP – Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior do Estado de São Paulo e os <b>AUXILIARES</b> DE ADMINISTRAÇÃO ESCOLAR, neste ato representado pelas ENTIDADES SINDICAIS PROFISSIONAIS, signatárias da Convenção Coletiva de Trabalho 2005-2006 a criação do BANCO DE HORAS.
<tr><td class=campo style="text-align:justify"><b>Cláusula Segunda</b> – A partir de 01 de março de 2005, fica instituído para a categoria dos <b>AUXILIARES</b> de Administração Escolar, o Sistema de Banco de Horas, com base na Lei 9.601, de 21-01-98, que deu nova redação ao § 2° do artigo 59 da Consolidação das Leis do Trabalho e a ele (art. 59) acrescentou o § 3°.
<br><b>§ 1º</b> – Será formado um banco, proveniente das horas trabalhadas além da jornada normal diária, as quais serão compensadas nos termos do presente Acordo.
<br><b>§ 2º</b> – A composição do banco de horas se dará mediante o acúmulo, apurado por meio de cartão de ponto, de horas credoras ou devedoras.
<br><b>§ 3º</b> – As horas excedentes, a que se refere o parágrafo 2°, estarão limitadas a 2 (duas) horas diárias e 10 (dez) horas semanais, as quais serão acumuladas para futura compensação.
<br><b>§ 4º</b> – Será permitido um saldo negativo de, no máximo, 30 horas a serem compensadas, conforme estabelecido nos parágrafos 6° a 12°.
<br><b>§ 5º</b> – As horas que ultrapassarem o limite estabelecido no parágrafo 3° desta cláusula serão remuneradas como horas extras, em conformidade com a cláusula 09 da Convenção Coletiva de Trabalho 2005.
<br><b>§ 6º</b> – A compensação não poderá ocorrer nas Férias, Feriados e Descanso Semanal Remunerado.
<br><b>§ 7º</b> – Sempre que houver interesse das partes em que haja a compensação, tal solicitação se dará com antecedência mínima de 48 (quarenta e oito) horas.
<br><b>§ 8º</b> – A cada 120 (cento e vinte) dias serão realizados balanços para apuração do saldo de horas e planejamento da compensação. Havendo interesse entre as partes, o saldo existente poderá ser transferido, todo ou em parte, para o balanço do período seguinte. Poderá, ainda, o saldo apurado ser remunerado como hora extra, conforme o disposto na cláusula 9 da Convenção Coletiva de Trabalho 2006/2006.
<br><b>§ 9º</b> – A apuração e compensação de saldo negativo obedecerá ao mesmo critério do parágrafo anterior.
<br><b>§ 10</b> – Os atrasos, saídas e faltas por motivo justificado e não previsto na legislação ou na CCT 2005/2006, poderão ser compensados no Banco de Horas, limitando-se em uma ocorrência por semana.
<br><b>§ 11</b> – Os <b>AUXILIARES</b> contratados por prazo determinado, bem como aqueles que estão em período de experiência, não poderão valer-se do sistema de Banco de Horas.
<br><b>§ 12</b> – Nos casos de desligamento de <b>AUXILIARES</b> durante a vigência deste Acordo, obrigar-se-á a <b>MANTENEDORA</b> a pagar o adicional previsto na cláusula 9ª da CCT 2005/2006, sobre as horas não compensadas, calculadas sobre o valor da remuneração na data da rescisão. Na existência de horas a compensar (saldo negativo), conforme previsto nos parágrafos 6° e 9°, estas serão descontadas das verbas rescisórias.
<br><b>§ 13</b> – Qualquer divergência na aplicação deste Acordo deverá ser resolvida através da convocação do Foro para Solução de Conflitos Coletivos, conforme a cláusula 37 da CCT 2005/2006.
<br><b>§ 14</b> – A renovação, alteração ou rescisão deste Acordo dependerá de acordo escrito dos representantes das partes, antes de expirado seu prazo de validade.
<br><b>§ 15</b> – O prazo de vigência desta cláusula é de 12 (doze) meses, encerrando-se em 28 de fevereiro de 2006.
</table> 

<DIV style="page-break-after:always"></DIV>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td class=titulo align="center">ANEXO 02
<tr><td class=titulo align="center">INSTRUMENTO DE ADITAMENTO DA CONVENÇÃO COLETIVA DE TRABALHO
<tr><td class=campo style="text-align:justify">REGULAMENTO DO NÚCLEO INTERSINDICAL DE CONCILIAÇÃO TRABALHISTA 

<tr><td class=campo style="text-align:justify">Regulamento para funcionamento do Núcleo Intersindical de Conciliação Trabalhista entre o Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de São Paulo - SEMESP e o Sindicato .........................................

<tr><td class=campo style="text-align:justify">Através do presente Instrumento de Aditamento, as partes dão cumprimento ao que foi estipulado no parágrafo primeiro da cláusula 53 da Convenção Coletiva de Trabalho firmada entre as MANTENEDORAS e os AUXILIARESDE ADMINISTRAÇÃO ESCOLAR, implementando a criação do Núcleo Intersindical de Conciliação Trabalhista previsto na Lei nº 9958/2000, tudo nos termos das seguintes cláusulas e condições que têm como certas e ajustadas.

<tr><td class=campo style="text-align:justify"><b>1.</b>
<tr><td class=campo style="text-align:justify">Fica criado o Núcleo Intersindical de Conciliação Trabalhista entre o Sindicato das Entidades Mantenedoras de Estabelecimentos de Ensino Superior no Estado de São Paulo - SEMESP e o Sindicato ......................................................................... previsto na cláusula 53 da Convenção Coletiva de Trabalho entre estas partes, bem como, no artigo 625-A da Consolidação das Leis do Trabalho.

<tr><td class=campo style="text-align:justify"><b>2.</b>
<tr><td class=campo style="text-align:justify">O Núcleo aqui mencionado irá funcionar na cidade de ..........................................................., à .....................................................

<tr><td class=campo style="text-align:justify"><b>3.</b>
<tr><td class=campo style="text-align:justify">Os trabalhos do Núcleo obedecerão ao presente Regulamento, aprovado pelos convenentes.

<tr><td class=campo style="text-align:justify"><b>4.</b>
<tr><td class=campo style="text-align:justify">O Núcleo Intersindical de Conciliação Trabalhista, doravante denominado simplesmente de Comissão, funcionará nos termos previstos na Lei 9958/2000, com a finalidade de servir de instrumento para rápida solução dos conflitos de trabalho.

<tr><td class=campo style="text-align:justify"><b>5.</b>
<tr><td class=campo style="text-align:justify">Para acionar os préstimos da Comissão, o interessado deverá protocolar na sede de funcionamento da comissão, pedido de intervenção conciliatória, em quatro vias, sendo uma para arquivo na Comissão, outra para a notificação da parte contrária e as restantes para as Entidades Sindicais signatárias.

<tr><td class=campo style="text-align:justify"><b>6.</b>
<tr><td class=campo style="text-align:justify">Tal pedido deverá expor de modo sintético os fatos e os fundamentos da questão, bem como, os valores pretendidos pelo interessado em razão de tal formulação.

<tr><td class=campo style="text-align:justify"><b>7.</b>
<tr><td class=campo style="text-align:justify">O interessado poderá fazer-se representar por advogado na apresentação do pedido inicial, bem como, fazer-se acompanhar de tal profissional quando da sessão de conciliação. Nesta oportunidade, a empresa deverá comparecer na pessoa de seu representante legal ou por preposto, com poderes específicos para transigir e firmar termo de conciliação.

<tr><td class=campo style="text-align:justify"><b>8.</b>
<tr><td class=campo style="text-align:justify">Recebido o pedido de intervenção conciliatória, a Comissão fixará de imediato, data e hora para a sessão de conciliação, saindo intimado o interessado e notificando-se a parte contrária por escrito. Tal deverá realizar-se no máximo em dez dias, a contar da data do protocolo.

<tr><td class=campo style="text-align:justify"><b>9.</b>
<tr><td class=campo style="text-align:justify">A conciliação praticada perante a Comissão, não poderá ser de caráter genérico, somente sendo admissível homologar transações sobre matéria constante do pedido inicial, conforme disposto na cláusula 6ª do presente instrumento. Será permitido aos interessados, inclusive, ressalvar expressamente que a transação não abrange alguma questão especificamente destacada.

<tr><td class=campo style="text-align:justify"><b>10.</b>
<tr><td class=campo style="text-align:justify">Aberta a sessão conciliatória, os membros da Comissão explicarão às partes presentes qual a natureza das funções do órgão, bem como, tecerão as ponderações necessárias à mediação para a solução negocial do conflito.

<tr><td class=campo style="text-align:justify"><b>11.</b>
<tr><td class=campo style="text-align:justify">Obtida ou não a conciliação entre as partes, será lavrado o termo respectivo para as finalidades previstas no parágrafo segundo do artigo 625-D ou no artigo 625-E da Lei 9958/2000.

<tr><td class=campo style="text-align:justify"><b>12.</b>
<tr><td class=campo style="text-align:justify">O Núcleo deverá intentar realizar a sessão de conciliação no prazo de 10 (dez) dias, a contar da provocação do interessado. Não se ultimando a tentativa em tal prazo, será fornecida certidão negativa ao interessado para os fins de Direito.

<tr><td class=campo style="text-align:justify"><b>13.</b>
<tr><td class=campo style="text-align:justify">Os trabalhos do Núcleo serão desenvolvidos por conciliadores indicados pela Entidades Sindicais signatárias, em número de 3 (três) para cada parte conveniente. Em cada sessão realizada, os interessados serão sempre atendidos por, pelo menos, dois conciliadores, sendo um representante da Entidade Sindical patronal e outro da entidade Sindical profissional.

<tr><td class=campo style="text-align:justify"><b>14.</b>
<tr><td class=campo style="text-align:justify">Para que produza seus efeitos jurídicos, assinaram o presente na forma da lei.

<tr><td class=campo style="text-align:justify">São Paulo, .... de junho de 2005 

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