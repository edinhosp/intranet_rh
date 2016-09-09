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
<title>Convenção Coletiva 2015/16 - Professores</title>
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
<tr><td class=titulo align="center">CONVENÇÃO COLETIVA DE TRABALHO PARA 2015/16
<tr><td class=titulo align="center">SEMESP
<tr><td class=titulo align="center">PROFESSORES - ENSINO SUPERIOR
<tr><td class=campo style="text-align:justify">

<tr><td class=titulo>1. Abrangência
<tr><td class=campo style="text-align:justify">Esta Convenção abrange a categoria econômica dos estabelecimentos particulares de ensino superior no Estado de São Paulo, aqui designados como MANTENEDORA e a categoria profissional diferenciada dos professores, aqui designada simplesmente como PROFESSOR.
<br><b>Parágrafo primeiro</b> – A categoria dos PROFESSORES abrange todos aqueles que exercem a atividade docente, independentemente da denominação sob a qual a função for exercida. Considera-se atividade docente a função de ministrar aula.
<br><b>Parágrafo segundo</b> – Quando o PROFESSR for contratado em um município para exercer a sua atividade em outro, prevalecerá o cumprimento da Convenção Coletiva do município em que o serviço é prestado

<tr><td class=titulo>2. Duração
<tr><td class=campo style="text-align:justify">Esta Convenção Coletiva de Trabalho terá duração de um ano, com vigência de 1º de março de 2015 a 29 de fevereiro de 2016.

<tr><td class="campot" align="center"><b>Salários, reajuste e pagamento
<tr><td class="campov" align="center"><b>Reajustes/Correções salariais

<tr><td class=titulo>3. Reajuste salarial em 2015
<tr><td class=campo style="text-align:justify">No ano de 2015 as MANTENEDORAS deverão aplicar os seguintes índices de reajuste sobre a remuneração mensal devida aos seus PROFESSORES em 1º de março de 2014:
<br>- 7,41% (sete virgula quarenta e um por cento), a partir de 1º de março;
<br>- 8,00% (oito por cento), a partir de 1º de julho.
<br><b>Parágrafo primeiro</b> – As diferenças salariais relativas aos meses de março, abril e maio de 2015 deverão ser pagas até o dia 12 de junho de 2015, sob pena de, em não o fazendo, arcar com a multa estabelecida na cláusula Prazo para pagamento de salários desta Convenção.
<br><b>Parágrafo segundo</b> – Fica estabelecido que a remuneração mensal de 1º de julho de 2015, reajustado pelo índice definido nesta cláusula, servirá como base de cálculo para a data base de 1º de março de 2016.

<tr><td class=titulo>4. Compensações salariais
<tr><td class=campo style="text-align:justify">No ano de 2015 será permitida a compensação de eventuais antecipações salariais concedidas no período compreendido entre 1º de março de 2014 e 28 de fevereiro de 2015.
<br><b>Parágrafo único</b> – Não será permitida a compensação daquelas antecipações salariais que decorrerem de promoções, transferências, ascensão em plano de carreira e os reajustes concedidos com cláusula expressa de não compensação.

<tr><td class="campov" align="center"><b>Pagamento de salário: formas e prazos

<tr><td class=titulo>5. Composição da remuneração mensal do professor
<tr><td class=campo style="text-align:justify">A remuneração mensal do PROFESSOR é composta, no mínimo, por três itens: o salário base, o descanso semanal remunerado (DSR) e a hora-atividade.
<br>O salário base é calculado pela seguinte equação: número de aulas semanais multiplicado por 4,5 semanas e multiplicado, ainda, pelo valor da hora-aula (artigo 320, parágrafo 1º da CLT).
<br>O DSR corresponde a 1/6 (um sexto) do salário base, acrescido, quando houver, do total de horas extras e do adicional noturno (Lei 605/49).
<br>A hora-atividade corresponde a 5% (cinco por cento) do total obtido com a somatória de todos os valores acima referidos.
<br><b>Parágrafo único</b> - A remuneração adicional do PROFESSOR pelo exercício concomitante de função não docente obedecerá aos critérios estabelecidos entre a MANTENEDORA e o PROFESSOR que aceitar o cargo.

<tr><td class=titulo>6. Prazo para pagamento de salários
<tr><td class=campo style="text-align:justify">Os salários deverão ser pagos, no máximo, até o quinto dia útil do mês subsequente ao trabalhado, considerando que sábado é dia útil, conforme Instrução Normativa número 01 do MTE, de 7/11/1989.
<br><b>Parágrafo único</b> - O não pagamento dos salários no prazo obriga a MANTENEDORA a pagar multa diária, em favor do PROFESSOR, no valor de 1/50 (um cinquenta avos) de seu salário mensal.

<tr><td class=titulo>7. Comprovante de pagamento
<tr><td class=campo style="text-align:justify">A MANTENEDORA deverá fornecer ao PROFESSOR, mensalmente, comprovante de pagamento, devendo estar discriminados: 
<blockquote style="margin-top:0;margin-bottom:0">
a) identificação da MANTENEDORA e do estabelecimento de ensino; 
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
<br>m) outros descontos.
</blockquote>

<tr><td class="campov" align="center"><b>Descontos salariais

<tr><td class=titulo>8. Autorização para desconto em folha de pagamento
<tr><td class=campo style="text-align:justify">O desconto do professor em folha de pagamento somente poderá ser realizado mediante sua autorização, nos termos dos artigos 462 e 545 da CLT, quando os valores forem destinados ao custeio de prêmios de seguro, planos de saúde, mensalidades associativas ou outras que constem da sua expressa autorização, desde que não haja previsão expressa de desconto na presente norma coletiva.
<br><b>Parágrafo único</b> – Encontra-se no Sindicato, à disposição da MANTENEDORA, devendo ser a ela encaminhada, quando solicitada formalmente, cópia de autorização do PROFESSOR para o desconto da mensalidade associativa.

<tr><td class="campot" align="center"><b>Gratificações, adicionais, auxílios e outros
<tr><td class="campov" align="center"><b>Adicional de hora extra

<tr><td class=titulo>9. Horas extras
<tr><td class=campo style="text-align:justify">Considera-se atividade extra todo trabalho desenvolvido em horário diferente daquele habitualmente realizado na semana. As atividades extras devem ser pagas com adicional de 100% (cem por cento).
<br><b>Parágrafo primeiro</b> – Não é considerada atividade extra a participação em cursos de capacitação e aperfeiçoamento docente, desde que aceita livremente pelo PROFESSOR.
<br><b>Parágrafo segundo</b> – Serão pagas apenas como aulas normais, acrescidas do DSR e da hora-atividade, aquelas que forem adicionadas provisoriamente à carga horária habitual, decorrentes:
<blockquote style="margin-top:0;margin-bottom:0">
a) da substituição temporária de outro PROFESSOR, com duração predeterminada, decorrente de licença médica, maternidade ou para estudos. Nestes casos, a substituição deverá ser formalizada através de documento firmado entre a MANTENEDORA e o PROFESSOR que aceitar realizá-la;
<br>b) de substituições eventuais de faltas de PROFESSOR responsável, desde que aceitas livremente pelo PROFESSOR substituto;
<br>c) de reposição de eventuais faltas que foram descontadas dos salários nos meses em que ocorreram;
<br>d) da realização de cursos eventuais ou de curta duração, inclusive cursos de dependência, e aceitas livremente, mediante documento firmado entre o PROFESSOR convidado a ministrá-los e a MANTENEDORA.
<br>e) do comparecimento a reuniões didático-pedagógicas, de avaliação e de planejamento, quando realizadas fora de seu horário habitual de trabalho, desde que aceito livremente pelo PROFESSOR.
</blockquote>
<b>Parágrafo terceiro</b> – A participação em Comissões Internas e Externas da Unidade de Ensino da MANTENEDORA, desde que aceita livremente pelo PROFESSOR mediante documento firmado, será remunerada como aula ou hora normal, acrescida de DSR.

<tr><td class="campov" align="center"><b>Adicional noturno

<tr><td class=titulo>10. Adicional noturno
<tr><td class=campo style="text-align:justify">O trabalho noturno deve ser pago nas atividades realizadas após as 22 (vinte e duas) horas e corresponde a 25% (vinte e cinco por cento) do valor da hora-aula.

<tr><td class="campov" align="center"><b>Outros adicionais

<tr><td class=titulo>11. Hora-atividade
<tr><td class=campo style="text-align:justify">Fica mantido o adicional de 5% (cinco por cento) a título de hora-atividade, destinado exclusivamente ao pagamento do tempo gasto pelo PROFESSOR, fora do estabelecimento de ensino, na preparação de aulas, provas e exercícios, bem como na correção dos mesmos.

<tr><td class=titulo>12. Adicional por atividades em outros municípios
<tr><td class=campo style="text-align:justify">Quando o PROFESSOR desenvolver suas atividades a serviço da mesma MANTENEDORA em município diferente daquele onde foi contratado e onde ocorre a prestação habitual do trabalho, deverá receber um adicional de 25% (vinte e cinco por cento) sobre o total de sua remuneração no novo município. Quando o PROFESSOR voltar a prestar serviços no município de origem, cessará a obrigação no pagamento do adicional.
<br><b>Parágrafo primeiro</b> - Nos casos em que ocorrer a transferência definitiva do PROFESSOR, aceita livremente por este, em documento firmado entre as partes, não haverá a incidência do adicional referido no caput, obrigando-se a MANTENEDORA a efetuar o pagamento de um único salário mensal integral, ao PROFESSOR, no ato da transferência, a título de ajuda de custo.
<br><b>Parágrafo segundo</b> - Fica assegurada a garantia de emprego pelo período de seis meses ao PROFESSOR transferido de município, contados a partir do início do trabalho e/ou da efetivação da transferência.
<br><b>Parágrafo terceiro</b> – Caso a MANTENEDORA desenvolva atividade acadêmica em municípios considerados conurbados, poderá solicitar isenção do pagamento do adicional determinado no caput, desde que encaminhe material comprobatório ao SEMESP, para análise e deliberação do Foro Conciliatório para Solução de Conflitos Coletivos, previsto na presente Convenção.

<tr><td class="campov" align="center"><b>Auxílio educação

<tr><td class=titulo>13. Bolsas de estudo
<tr><td class=campo style="text-align:justify"><b>A - Programa de Capacitação do Professor</b>
<tr><td class=campo style="text-align:justify">Todo PROFESSOR tem direito a bolsa de estudo integral, incluindo matrícula, em cursos de graduação, sequenciais e pós-graduação existentes e administrados pela MANTENEDORA que o emprega, observado o que segue:
<blockquote style="margin-top:0;margin-bottom:0">
1. A MANTENEDORA está obrigada a conceder, no máximo, duas bolsas de estudo, sendo que, nos cursos de graduação e sequenciais, não será possível que o PROFESSOR conclua mais de um curso nessa condição.
<br>2. As bolsas de estudo integrais em cursos de pós-graduação ou especialização existentes e administrados pela MANTENEDORA são válidas exclusivamente para o PROFESSOR, em áreas correlatas às disciplinas que o mesmo ministra na Instituição e que visem a capacitação docente, respeitados os critérios de seleção exigidos para ingresso no mesmo e obedecerão as seguintes condições :
<blockquote style="margin-top:0;margin-bottom:0">
	a) nos cursos stricto sensu ou de especialização que fixem um número máximo de alunos por turma, são limitadas em 30% (trinta por cento) do total de vagas oferecidas;
	<br>b) nos cursos de pós-graduação lato sensu não haverá limites de vagas. Caso a estrutura do curso torne necessária a limitação do número de alunos será observado o disposto na alínea “a” deste item.
</blockquote>	
3. O direito às bolsas de estudo passa a vigorar ao término do contrato de experiência, cuja duração não pode exceder de 90 (noventa) dias, conforme parágrafo único do artigo 445 da CLT.
<br>4. As bolsas de estudo serão mantidas quando o PROFESSOR estiver licenciado para tratamento de saúde ou em gozo de licença mediante anuência da MANTENEDORA, excetuado o disposto na cláusula “Licença sem Remuneração”.
<br>5. O PROFESSOR que for reprovado no período letivo perderá o direito à bolsa de estudo, voltando a gozar do benefício quando lograr aprovação no referido período. As disciplinas cursadas em regime de dependência serão de total responsabilidade do PROFESSOR, arcando o mesmo com o seu custo.
</blockquote>
<tr><td class=campo style="text-align:justify"><b>B - Programa de Inclusão, Capacitação para Filhos, Dependentes Legais e Estudantes
<tr><td class=campo style="text-align:justify">O CEBRADE – Centro Brasileiro de Desenvolvimento do Ensino Superior – tem, como um dos seus objetivos, desenvolver o Programa de Amparo Educativo Temporário – PAET, concedendo bolsas de estudo em Instituições Privadas de Ensino Superior. Os filhos ou dependentes legais do PROFESSOR têm direito a usufruir as gratuidades integrais do PAET, sem qualquer ônus, nos cursos de graduação ou sequenciais existentes e administrados pela MANTENEDORA para a qual o PROFESSOR trabalha, observado o disposto nesta cláusula e no “Regulamento do Programa de Capacitação”, anexado à presente Convenção.
<br><b>Parágrafo primeiro</b> – A MANTENEDORA deverá disponibilizar ao CEBRADE, mediante requerimento, bolsas de estudo em número suficiente para o atendimento da concessão das gratuidades integrais do PAET nas Instituições de Ensino Superior por ela mantida, para filhos ou dependentes legais dos seus PROFESSORES, observada a limitação de duas bolsas de estudo por PROFESSOR.
<br><b>Parágrafo segundo</b> – O beneficiário bolsista, concluinte de curso de graduação ou seqüencial, não poderá obter nova concessão de gratuidade em um desses cursos, na mesma IES.
<br><b>Parágrafo terceiro</b> – O SEMESP e a FEDERAÇÃO representante da categoria profissional fiscalizarão o CEBRADE na gestão do Programa de Amparo Educativo Temporário para os filhos e dependentes legais dos PROFESSORES, na conformidade do estabelecido nesta cláusula e no “Regulamento do Programa de Capacitação”.
<br><b>Parágrafo quarto</b> – Para a concessão das gratuidades integrais aos filhos e dependentes legais do PROFESSOR, o CEBRADE não poderá fazer qualquer outra exigência a não ser o comprovante de aprovação no processo seletivo da IES administrado pela MANTENEDORA empregadora e a observância dos preceitos estabelecidos nesta cláusula e no “Regulamento do Programa de Capacitação”.
<br><b>Parágrafo quinto</b> – Terão direito a requerer e obter do CEBRADE a concessão de bolsas integrais de estudo, os dependentes legais do PROFESSOR reconhecidos pela Legislação do Imposto de Renda, ou que estejam sob a sua guarda judicial e vivam sob sua dependência econômica, devidamente comprovada.
<br><b>Parágrafo sexto</b> – Os filhos do PROFESSOR terão direito a obter do CEBRADE a concessão de bolsas de estudo integrais, sem qualquer ônus, desde que não tenham 25 (vinte e cinco) anos completos ou mais na data da efetivação da matrícula no curso superior. Os filhos ou dependentes legais do PROFESSOR serão denominados dependentes beneficiários.
<br><b>Parágrafo sétimo</b> – As gratuidades integrais serão mantidas aos dependentes beneficiários quando o PROFESSOR estiver licenciado para tratamento de saúde ou mediante anuência da MANTENEDORA, excetuado o disposto na cláusula “Licença sem remuneração” da presente Convenção.
<br><b>Parágrafo oitavo</b> – No caso de falecimento do PROFESSOR, os dependentes beneficiários continuarão a usufruir as gratuidades integrais até o final do curso, arcando tão somente com as disciplinas cursadas em regime de dependência.
<br><b>Parágrafo nono</b> – No caso de dispensa imotivada do PROFESSOR, os dependentes beneficiários continuarão a usufruir as gratuidades integrais até o final do ano letivo, arcando tão somente com as disciplinas cursadas em regime de dependência.
<br><b>Parágrafo décimo</b> – Os dependentes beneficiários que forem reprovados no período letivo perderão o direito à bolsa de estudo, voltando a gozar do benefício quando lograrem aprovação naquele período. As disciplinas cursadas em regime de dependência serão de total responsabilidade dos dependentes beneficiários, que deverão arcar com seu custo.
<br><b>Parágrafo onze</b> – Para usufruir as gratuidades integrais dos dependentes beneficiários, não se poderá exigir do PROFESSOR pagamento algum, a qualquer título, nem mesmo condicionar a concessão do benefício à associação, sindicalização ou filiação.
<br><b>Parágrafo doze</b> – Caso a MANTENEDORA não queira participar do Programa de Amparo Educativo Temporário – PAET, gerenciado pelo CEBRADE, estará obrigada a conceder bolsas de estudo aos PROFESSORES que trabalham nas Instituições de Ensino Superior por elas mantidas ou administradas, nas condições e termos estabelecidos nesta cláusula e no Regulamento em anexo.
<br><b>Parágrafo treze</b> – Além dos casos previstos nesta cláusula, a MANTENEDORA poderá fornecer outras bolsas de estudos, cujas condições serão objeto de termo aditivo a ser firmado entre MANTENEDORA e CEBRADE.

<tr><td class="campov" align="center"><b>Auxílio-saúde

<tr><td class=titulo>14. Assistência médico-hospitalar
<tr><td class=campo style="text-align:justify">A MANTENEDORA está obrigada a assegurar, às suas expensas, nos limites estabelecidos nesta cláusula, assistência médico-hospitalar a todos os seus PROFESSORES, sendo-lhe facultada a escolha por plano de saúde, seguro-saúde ou convênios com empresas prestadoras de serviços médico-hospitalares. Poderá ainda prestar a referida assistência diretamente, em se tratando de instituições que disponham de serviços de saúde e hospitais próprios ou conveniados. Qualquer que seja a opção feita, a assistência médico-hospitalar deve assegurar as condições e os requisitos mínimos que seguem relacionados:
<br><b>1. Abrangência</b>
<blockquote style="margin-top:0;margin-bottom:0">
A assistência médico-hospitalar deve ser realizada no município onde funciona o estabelecimento de ensino superior ou onde vive o PROFESSOR, a critério da MANTENEDORA. Em casos de emergência, deverá haver garantia de atendimento integral em qualquer localidade do Estado de São Paulo ou fixação, em contrato, de formas de reembolso.
</blockquote>
<b>2. Coberturas mínimas</b>
<blockquote style="margin-top:0;margin-bottom:0">
2.1 Quarto para quatro pacientes, no máximo.
<br>2.2 Consultas.
<br>2.3 Prazo de internação de 365 (trezentos e sessenta e cinco) dias por ano (comum e UTI/CTI)
<br>2.4 Parto, independentemente do estado gravídico.
<br>2.5 Moléstias infecto-contagiosas que exijam internação.
<br>2.6 Exames laboratoriais, ambulatoriais e hospitalares.
</blockquote>
<b>3. Carência</b>
<blockquote style="margin-top:0;margin-bottom:0">
Não haverá carência na prestação dos serviços médicos e laboratoriais.
</blockquote>
<b>4. Professor ingressante</b>
<blockquote style="margin-top:0;margin-bottom:0">
Não haverá carência para o PROFESSOR ingressante, independentemente do mês em que for contratado.
</blockquote>
<b>5. Pagamento</b>
<blockquote style="margin-top:0;margin-bottom:0">
Caberá ao PROFESSOR o pagamento de 10% (dez por cento) do valor da Assistência Médica, respeitado o disposto nos parágrafos 1º, 2º e 3º.
</blockquote>
	<b>Parágrafo primeiro</b> – A MANTENEDORA deverá enviar ao Sindicato cópia do contrato formalizado com a empresa de assistência médico–hospitalar ou de seguro saúde ou de medicina de grupo que comprove o valor pago.
<br><b>Parágrafo segundo</b> – Caso a assistência médico-hospitalar vigente na Instituição venha a sofrer reajuste em virtude de possíveis modificações estabelecidas em legislação que abranja o segmento - Lei 9.656, de 03 de junho de 1998 e MP 2.097-39, de 26 de abril de 2001, ou que vierem a ser estabelecidas em lei, ou por mudança de empresa prestadora de serviço, a pedido dos empregados da Instituição ou por quebra de contrato, unilateralmente, por parte da atual empresa prestadora de serviço, a MANTENEDORA continuará a contribuir com o valor mensal vigente até a data da modificação, devendo o PROFESSOR arcar com o valor excedente, que será descontado em folha e consignado no comprovante de pagamento, nos termos do artigo 462 da CLT.
<br><b>Parágrafo terceiro</b> – Caso ocorra mudança de empresa prestadora de serviço, por decisão unilateral da MANTENEDORA, com conseqüente reajuste no valor vigente, o PROFESSOR estará isento do pagamento do valor excedente, cabendo à MANTENEDORA prover integralmente a assistência médico-hospitalar, sem nenhum ônus para o PROFESSOR.
<br><b>Parágrafo quarto</b> – Para efeito do disposto no parágrafo primeiro desta cláusula, caberá à MANTENEDORA remeter a documentação comprobatória para análise e deliberação da Comissão Permanente de Negociação.
<br><b>Parágrafo quinto</b> – Fica facultado ao PROFESSOR optar pela prestação de assistência médico-hospitalar em uma única instituição de ensino, quando mantiver mais de um vínculo empregatício como PROFESSOR. É necessário que o PROFESSOR se manifeste por escrito, com antecedência mínima de vinte dias, para que a MANTENEDORA possa proceder à suspensão dos serviços.
<br><b>Parágrafo sexto</b> – Caso o PROFESSOR mantenha vínculo empregatício com mais de uma Instituição de Ensino, as MANTENEDORAS, em conjunto, poderão optar por conceder-lhe um único plano de saúde, pago por elas, em regime de cotização de custos, respeitadas as condições estabelecidas nesta cláusula.
<br><b>Parágrafo sétimo</b> – Mediante pagamento complementar e adesão facultativa, devidamente documentada, o PROFESSOR poderá optar pela ampliação dos serviços de saúde garantidos nesta Convenção ou estendê-los a seus dependentes.

<tr><td class="campov" align="center"><b>Auxílio-creche

<tr><td class=titulo>15. Creches
<tr><td class=campo style="text-align:justify">É obrigatória a instalação de local destinado a guarda de crianças de até seis meses, quando a MANTENEDORA mantiver contratada, em jornada integral, pelo menos trinta funcionárias com idade superior a 16 anos. A manutenção da creche poderá ser substituída pelo pagamento do reembolso-creche, nos termos da legislação em vigor (artigo 389, parágrafo 1º da CLT e Portarias MTE nº 3296 de 3/9/1986 e nº 670 de 27/8/1997) ou, ainda, a celebração de convênio com entidade de idoneidade reconhecida.

<tr><td class="campot" align="center"><b>Contrato de trabalho: admissão, demissão, modalidades
<tr><td class="campov" align="center"><b>Normas para admissão/contratação

<tr><td class=titulo>16. Remuneração mensal ou valor da hora aula do PROFESSOR ingressante na MANTENEDORA
<tr><td class=campo style="text-align:justify">A MANTENEDORA não poderá contratar PROFESSOR cuja remuneração mensal ou o valor da hora aula seja inferior ao valor da remuneração mensal ou da hora aula mínima dos PROFESSORES mais antigos que possuam o mesmo grau de qualificação ou titulação de quem está sendo contratado, respeitado o quadro de carreira da MANTENEDORA.
<br><b>Parágrafo único</b> – Ao PROFESSOR admitido após 1º de março de 2015 serão concedidos os mesmos percentuais de reajustes e aumentos salariais estabelecidos na cláusula Reajuste salarial em 1º de março de 2015.

<tr><td class=titulo>17. Readmissão do professor
<tr><td class=campo style="text-align:justify">O PROFESSOR que for readmitido até doze meses após o seu desligamento ficará desobrigado de firmar contrato de experiência.

<tr><td class=titulo>18. Anotações na carteira de trabalho
<tr><td class=campo style="text-align:justify">A MANTENEDORA está obrigada a promover, em quarenta e oito horas, as anotações nas Carteiras de Trabalho de seus PROFESSORES, ressalvados eventuais prazos mais amplos permitidos por lei.
<br><b>Parágrafo único</b> – É obrigatória a anotação na Carteira de Trabalho das mudanças provocadas por ascensão ou alteração de titulação, decorrentes e previstas em plano de carreira.

<tr><td class="campov" align="center"><b>Desligamento / demissão

<tr><td class=titulo>19. Garantia semestral de salários
<tr><td class=campo style="text-align:justify">Ao PROFESSOR demitido sem justa causa, a MANTENEDORA garantirá:
<blockquote style="margin-top:0;margin-bottom:0">
a) no primeiro semestre, a partir de 1º de janeiro, as remunerações mensais integrais até o dia 30 de junho;
<br>b) no segundo semestre, as remunerações mensais integrais até o dia 31 de dezembro, ressalvado o parágrafo 4º.
</blockquote>
	<b>Parágrafo primeiro</b> - Não terá direito à Garantia Semestral de Salários o PROFESSOR que, na data da comunicação da dispensa, contar com menos de 18 (dezoito) meses de serviço prestado à MANTENEDORA, ressalvado o parágrafo 4º desta cláusula.
<br><b>Parágrafo segundo</b> – No caso de demissões efetuadas no final do primeiro semestre letivo, para não ficar obrigada a pagar ao PROFESSOR os salários do segundo semestre, a MANTENEDORA deverá observar as seguintes disposições:
<blockquote style="margin-top:0;margin-bottom:0">
	a) com aviso prévio a ser trabalhado, a demissão deverá ser formalizada com antecedência mínima de trinta dias do início das férias;
	<br>b) sendo o aviso prévio indenizado, a demissão deverá ser formalizada até um dia antes do início das férias, ainda que as férias tenham seu início programado para o mês de julho, obedecendo ao que dispõe a cláusula “Férias” da presente Convenção.
</blockquote>
	<b>Parágrafo terceiro</b> - No caso de demissões efetuadas no final do ano letivo, para não ficar obrigada a pagar ao PROFESSOR os salários do primeiro semestre do ano seguinte, a MANTENEDORA deverá observar as seguintes disposições:
<blockquote style="margin-top:0;margin-bottom:0">	
	a) com aviso prévio a ser trabalhado, a demissão deverá ser formalizada com antecedência mínima de trinta dias do início do recesso escolar;
	<br>b) sendo o aviso prévio indenizado, a demissão deverá ser formalizada até um dia antes do início do recesso escolar.
</blockquote>	
	<b>Parágrafo quarto</b> - Quando as demissões ocorrerem a partir de 16 de outubro, a MANTENEDORA pagará, independentemente do tempo de serviço do PROFESSOR, valor correspondente à remuneração devida até o dia 18 de janeiro, inclusive, do ano subsequente, respeitado o pagamento mínimo de 30 (trinta) dias, a título de férias escolares, para efeito do que define a súmula 10 do egrégio TST, ressalvados os contratos de experiência e por prazo determinado, estes últimos válidos somente nos casos de substituição temporária, conforme o disposto na alínea a) do parágrafo 2º da cláusula Horas extras da presente Convenção.
<br><b>Parágrafo quinto</b> – Na vigência da presente Convenção os PROFESSORES serão remunerados a partir da data de início de suas atividades na MANTENEDORA, incluindo o período de planejamento escolar.
<br><b>Parágrafo sexto</b> - As remunerações complementares previstas nesta cláusula terão natureza indenizatória, não integrando, para nenhum efeito legal, o tempo de serviço do PROFESSOR.

<tr><td class=titulo>20. Indenizações por dispensa imotivada
<tr><td class=campo style="text-align:justify">“O PROFESSOR demitido sem justa causa, além das indenizações previstas na cláusula “Garantia Semestral de Salários” desta Convenção, terá direito a receber o valor equivalente a 3 (três) dias para cada ano trabalhado na MANTENEDORA, nos termos da Lei nº 12.506/2011, sem o limite de tempo de serviço estabelecido na mesma, ressaltando que não há cumulatividade entre a lei e a previsão contida nesta norma coletiva.
<br><b>Parágrafo primeiro</b> – Caso o PROFESSOR tenha, à data do desligamento, no mínimo cinquenta anos de idade e conte com pelo menos um ano de serviço na MANTENEDORA, terá direito ainda a receber aviso prévio adicional indenizado de 15 (quinze) dias.
<br><b>Parágrafo segundo</b> – Não terá direito à indenização assegurada no parágrafo primeiro o PROFESSOR que na data de admissão na MANTENEDORA contar com mais de cinquenta anos de idade.
<br><b>Parágrafo terceiro</b> – O aviso-prévio, quando trabalhado, será de trinta dias, com as reduções previstas no artigo 488 da CLT. O adicional de três dias por ano trabalhado, na forma do caput, será sempre indenizado na rescisão contratual.

<tr><td class=titulo>21. Pedido de demissão no final de ano letivo
<tr><td class=campo style="text-align:justify">O PROFESSOR que no final do ano letivo comunicar sua demissão até o dia que antecede o início do recesso escolar, será dispensado do cumprimento do aviso prévio e terá direito a receber, como indenização, a remuneração até o dia 18 de janeiro do ano subsequente, independentemente do tempo de serviço na MANTENEDORA.

<tr><td class=titulo>22. Demissão por justa causa
<tr><td class=campo style="text-align:justify">Quando houver demissão por justa causa, nos termos do art. 482 da CLT, a MANTENEDORA está obrigada a determinar na carta-aviso o motivo que deu origem à dispensa. Caso contrário, fica descaracterizada a justa causa.

<tr><td class="campov" align="center"><b>Outras normas referentes à admissão, demissão e modalidades de contratação

<tr><td class=titulo>23. Multa por atraso na homologação.
<tr><td class=campo style="text-align:justify">A MANTENEDORA deve pagar as verbas devidas na rescisão contratual no dia seguinte ao término do aviso prévio, quando trabalhado, ou dez dias após o desligamento, quando houver dispensa do cumprimento de aviso prévio. O atraso no pagamento das verbas rescisórias obrigará a MANTENEDORA ao pagamento de multa, em favor do PROFESSOR, correspondente a um mês de sua remuneração, conforme o disposto no parágrafo 8º do artigo 477 da CLT.
<br>A partir do vigésimo dia de atraso da homologação da rescisão, a contar da data estabelecida pela legislação para o pagamento das verbas rescisórias, a MATENEDORA estará obrigada, ainda, a pagar ao PROFESSOR multa diária de 0,2% (dois décimos percentuais) do salário mensal.
<br><b>Parágrafo primeiro</b> - A MANTENEDORA deverá agendar a homologação no respectivo Sindicato no prazo máximo de dez dias após a dispensa do PROFESSOR e estará desobrigada de pagar a multa definida no caput, quando o atraso vier a ocorrer, comprovadamente, por motivos alheios à sua vontade.
<br><b>Parágrafo segundo</b> – O Sindicato está obrigado a fornecer comprovante de comparecimento sempre que a MANTENEDORA se apresentar para homologação das rescisões contratuais e comprovar a convocação do PROFESSOR.
<br><b>Parágrafo terceiro</b> – Nos termos da orientação jurisprudencial 82 do TST e da Instrução Normativa 15, de 14 de julho de 2010 do MTE, no que tange à anotação e baixa em CTPS quando o aviso prévio for indenizado, deverá ser anotado na página relativa ao contrato de trabalho, o último dia do aviso prévio projetado e na página de “anotações gerais” o último dia efetivamente trabalhado, consignando em TRCT a data de afastamento como a do último dia efetivamente trabalhado.

<tr><td class=titulo>24. Atestados de afastamento e salários
<tr><td class=campo style="text-align:justify">Sempre que solicitada, a MANTENEDORA deverá fornecer ao PROFESSOR atestado de afastamento e salário (AAS), previsto na legislação previdenciária.

<tr><td class="campot" align="center"><b>Relações de trabalho: duração, distribuição, controle, faltas
<tr><td class="campov" align="center"><b>Estabilidade mãe

<tr><td class=titulo>25. Garantia de emprego à gestante.
<tr><td class=campo style="text-align:justify">É proibida a dispensa arbitrária ou sem justa causa da PROFESSORA gestante, desde o início da gravidez até sessenta dias após o término do afastamento legal. O aviso-prévio começará a contar a partir do término do período de estabilidade.

<tr><td class="campov" align="center"><b>Estabilidade acidentados / portadores doença profissional

<tr><td class=titulo>26. Garantias ao professor com sequelas ocasionadas por doenças profissionais ou acidente de trabalho
<tr><td class=campo style="text-align:justify">Será garantida ao PROFESSOR acidentado no trabalho ou acometido por doença profissional a permanência na empresa em função compatível com o seu estado físico, sem prejuízo na remuneração antes percebida, desde que, após o acidente ou comprovação da aquisição de doença profissional, apresente, cumulativamente, redução da capacidade laboral, atestada pelo órgão oficial e que se tenha tornado incapaz de exercer a função que anteriormente desempenhava. Nessa situação, o PROFESSOR estará obrigado a participar dos processos de readaptação e reabilitação profissional.
<br><b>Parágrafo único</b> – O período de estabilidade do PROFESSOR que estiver participando de processos de readaptação e reabilitação profissional será o previsto em lei.

<tr><td class="campov" align="center"><b>Estabilidade portadores doença não profissional

<tr><td class=titulo>27. Estabilidade para portadores de doenças graves
<tr><td class=campo style="text-align:justify">Fica assegurada, até alta médica, considerada como apto ao trabalho, ou eventual concessão de aposentadoria por invalidez, estabilidade no emprego aos PROFESSORES acometidos por doenças graves ou incuráveis e aos PROFESSORES portadores do vírus HIV que vierem a apresentar qualquer tipo de infecção ou doença oportunista, resultante da patologia de base.
<br><b>Parágrafo único</b> – São consideradas doenças graves ou incuráveis, a tuberculose ativa, alienação mental, esclerose múltipla, neoplasia maligna, cegueira definitiva, hanseníase, cardiopatia grave, doença de Parkinson, paralisia irreversível e incapacitante, espondiloastrose anquilosante, neofropatia grave, estados do Mal de Paget (osteíte deformante) e contaminação grave por radiação.

<tr><td class="campov" align="center"><b>Estabilidade aposentadoria

<tr><td class=titulo>28. Garantias ao professor em vias de aposentadoria
<tr><td class=campo style="text-align:justify">Fica assegurado ao PROFESSOR que comprovadamente estiver a vinte e quatro meses ou menos da aposentadoria integral por tempo de serviço ou da aposentadoria por idade, a garantia de emprego durante o período que faltar até a aquisição do direito.
<br><b>Parágrafo primeiro</b> – A garantia de emprego é devida ao PROFESSOR que estiver contratado pela MANTENEDORA há pelo menos três anos.
<br><b>Parágrafo segundo</b> – A comprovação à MANTENEDORA deverá ser feita mediante a apresentação de documento que ateste o tempo de serviço. Este documento deverá ser emitido por pessoa credenciada junto ao órgão previdenciário. Se o PROFESSOR depender de documentação para realização da contagem, terá um prazo de trinta dias, a contar da data prevista ou marcada para homologação da rescisão contratual. Comprovada a solicitação de tal documentação, os prazos serão prorrogados até que a mesma seja emitida, assegurando-se, nessa situação, o pagamento dos salários pelo prazo máximo de cento e vinte dias.
<br><b>Parágrafo terceiro</b> – O contrato de trabalho do PROFESSOR só poderá ser rescindido por mútuo acordo homologado pelo Sindicato ou pedido de demissão.
<br><b>Parágrafo quarto</b> – Havendo acordo formal entre as partes, o PROFESSOR poderá exercer outra função, inerente ao magistério, durante o período em que estiver garantido pela estabilidade.
<br><b>Parágrafo quinto</b> – O aviso prévio, em caso de demissão sem justa causa, integra o período de estabilidade previsto nesta cláusula.
<br><b>Parágrafo sexto</b> – Para garantir a estabilidade prevista nesta cláusula, o PROFESSOR deverá encaminhar à MANTENEDORA, dentro da prorrogação prevista no parágrafo 2º, documentação que demonstre a tramitação do processo que atesta o tempo de serviço.

<tr><td class="campov" align="center"><b>Estabilidade adoção

<tr><td class=titulo>29. Licença por adoção ou guarda 
<tr><td class=campo style="text-align:justify">Nos termos da Lei 12.873, de 25/10/2013, será assegurada licença de 120 (cento e vinte) dias à PROFESSORA ou PROFESSOR que vier a adotar ou obtiver guarda judicial de crianças e fizer jus ao salário maternidade pago pela Previdência Social. 
<br><b>Parágrafo primeiro</b> – Não poderá ser concedido benefício a mais de um empregado, decorrente do mesmo processo de adoção ou guarda, ainda que cônjuges ou companheiros que estejam submetidos ao regime próprio da Previdência Social. 
<br><b>Parágrafo segundo</b> – Fica garantida a estabilidade no emprego ao PROFESSOR ou à PROFESSORA adotante, durante a licença e até 60 (sessenta) dias após o término do afastamento legal. O aviso-prévio começará a contar a partir do término do período de estabilidade.

<tr><td class="campov" align="center"><b>Outras normas de pessoal

<tr><td class=titulo>30. Mudança de disciplina
<tr><td class=campo style="text-align:justify">O PROFESSOR não poderá ser transferido de uma disciplina para outra, salvo com seu consentimento expresso e por escrito, sob pena de nulidade da referida transferência.

<tr><td class="campot" align="center"><b>Jornada de trabalho: duração, distribuição, controle, faltas
<tr><td class="campov" align="center"><b>Duração e horário

<tr><td class=titulo>31. Duração da hora-aula
<tr><td class=campo style="text-align:justify">A duração da hora-aula poderá ser de, no máximo, cinquenta minutos.
<br><b>Parágrafo primeiro</b> – Como exceção ao disposto no caput, a hora-aula poderá ter a duração de sessenta minutos nos cursos tecnológicos, desde que tenham sido autorizados ou reconhecidos com essa determinação expressa e cujos PROFESSORES desses cursos tenham sido contratados nessa condição.
<br><b>Parágrafo segundo</b> – As MANTENEDORAS de Instituições de Ensino que possuem cursos tecnológicos nas condições definidas no parágrafo 1º desta cláusula deverão apresentar à Comissão Permanente de Negociação definida na presente Convenção, até o dia 15 de agosto de 2015, a documentação de autorização ou reconhecimento do curso com a determinação expressa de hora-aula com duração de 60 (sessenta) minutos sob pena de, em não o fazendo, estar sujeita à majoração do valor do salário-aula de acordo com o que estabelece o parágrafo quarto desta cláusula.
<br><b>Parágrafo terceiro</b> – Caso a Comissão Permanente de Negociação delibere não ter havido determinação expressa do Ministério da Educação para que a duração da hora-aula dos cursos tecnológicos seja de 60 (sessenta) minutos, a MANTENEDORA deverá majorar o salário-aula de acordo com o que estabelece o parágrafo quarto desta cláusula.
<br><b>Parágrafo quarto</b> – Em caso de ampliação da duração da hora-aula vigente, respeitado o limite previsto no caput desta cláusula, a MANTENEDORA deverá acrescer ao salário-aula já pago, valor proporcional ao acréscimo do trabalho.

<tr><td class=titulo>32. Carga horária
<tr><td class=campo style="text-align:justify">Quando a MANTENEDORA e o PROFESSOR contratarem carga diária de aulas superior aos limites previstos no artigo 318 da CLT, o excedente à carga horária legal será remunerado como aula normal, acrescido de DSR, hora-atividade e vantagens pessoais.
<br><b>Parágrafo único</b> – Poderá ser flexibilizada a carga horária do PROFESSOR entre jornadas no exercício concomitante de função docente e atividade administrativa, não havendo assim pagamento, no intervalo, de horas aulas e salários, se o professor não tiver trabalhado no referido intervalo.

<tr><td class="campov" align="center"><b>Prorrogação / redução de jornada

<tr><td class=titulo>33. Irredutibilidade de carga horária e de remuneração
<tr><td class=campo style="text-align:justify">É proibida a redução de remuneração mensal ou de carga horária, ressalvada a ocorrência do disposto nas cláusulas Redução de carga horária por extinção de disciplina classe ou turma e Redução de carga horária por diminuição do número de alunos matriculados da presente Convenção, ou ainda, quando ocorrer iniciativa expressa do PROFESSOR. Em qualquer hipótese, é obrigatória a concordância recíproca, firmada por escrito.
<br><b>Parágrafo primeiro</b> – Não havendo concordância recíproca, a parte que deu origem à redução prevista nesta cláusula arcará com a responsabilidade da rescisão contratual.
<br><b>Parágrafo segundo</b> – Atividades administrativas, não inerentes ao trabalho docente, de duração temporária e determinada, poderão ser regulamentadas por contrato entre as partes, contendo a caracterização da atividade, o início e a previsão do término.
<br><b>Parágrafo terceiro</b> – A MANTENEDORA não poderá reduzir o valor da hora-aula dos contratos de trabalho vigentes, ainda que venha a instituir ou modificar plano de carreira.

<tr><td class=titulo>34. Redução de carga horária por extinção ou supressão de disciplina, classe ou turma
<tr><td class=campo style="text-align:justify">Ocorrendo supressão de disciplina, classe ou turma, em virtude de alteração na estrutura curricular prevista ou autorizada pela legislação vigente ou por dispositivo regimental devidamente aprovado por órgão colegiado da Instituição de Ensino, o PROFESSOR da disciplina, classe ou turma deverá ser comunicado da redução da sua carga horária, por escrito, com antecedência mínima de 30 (trinta) dias do início do período letivo e terá prioridade para preenchimento de vaga existente em outra classe ou turma ou em outra disciplina para a qual possua habilitação legal.
<br><b>Parágrafo primeiro</b> – O PROFESSOR deverá manifestar por escrito, no prazo máximo de 5 (cinco) dias após a comunicação da MANTENEDORA, a não-aceitação da transferência de disciplina ou de classe ou turma ou da redução parcial de sua carga horária. A ausência de manifestação do PROFESSOR caracterizará a sua aceitação.
<br><b>Parágrafo segundo</b> – Caso o PROFESSOR não aceite a transferência para outra disciplina, classe ou turma ou a redução parcial de carga horária, a MANTENEDORA deverá manter a carga horária semanal existente ou proceder à rescisão do contrato de trabalho, por demissão sem justa causa.

<tr><td class=titulo>35. Redução de carga horária por diminuição do número de alunos matriculados
<tr><td class=campo style="text-align:justify">Na ocorrência de diminuição do número de alunos matriculados que venha a caracterizar a supressão de turmas, curso ou disciplina, o PROFESSOR do curso em questão deverá ser comunicado, por escrito, da redução parcial ou total de sua carga horária no período compreendido entre o primeiro dia de aula e o último dia da segunda semana de aula do período letivo.
<br><b>Parágrafo primeiro</b> - O PROFESSOR deverá manifestar, também por escrito, a aceitação ou não da redução parcial de carga horária no prazo máximo de cinco dias após a comunicação da MANTENEDORA. A ausência de manifestação do PROFESSOR caracterizará a sua não aceitação.
<br><b>Parágrafo segundo</b> - Caso o PROFESSOR aceite a redução parcial de carga horária, deverá formalizar documento junto à MANTENEDORA e, em não aceitando, a MANTENEDORA deverá proceder à rescisão do contrato de trabalho, por demissão sem justa causa.
<br><b>Parágrafo terceiro</b> - Na hipótese de rescisão contratual, por demissão sem justa causa, o aviso prévio será indenizado, estando a MANTENEDORA desobrigada do pagamento do disposto na cláusula Garantia Semestral de Salários da presente Convenção.
<br><b>Parágrafo quarto</b> - Não ocorrendo redução do número de alunos matriculados que venha a caracterizar supressão do curso, de turma ou de disciplina, a MANTENEDORA que reduzir a carga horária do PROFESSOR estará sujeita ao disposto na cláusula “Garantia Semestral de Salários” desta Convenção quando ocorrer a rescisão do contrato de trabalho do PROFESSOR.

<tr><td class="campov" align="center"><b>Faltas

<tr><td class=titulo>36. Desconto de faltas
<tr><td class=campo style="text-align:justify">Na ocorrência de faltas, a MANTENEDORA poderá descontar da remuneração mensal do PROFESSOR, no máximo, o número de aulas em que o mesmo esteve ausente, o DSR (1/6), a hora-atividade e demais vantagens pessoais proporcionais a estas aulas.
<br><b>Parágrafo único</b> - É da competência e de integral responsabilidade da MANTENEDORA estabelecer mecanismos de controle de faltas e de pontualidade dos PROFESSORES, conforme a legislação vigente.

<tr><td class=titulo>37. Abono de faltas por casamento ou luto
<tr><td class=campo style="text-align:justify">Não serão descontadas, no curso de nove dias corridos, as faltas do PROFESSOR, por motivo de gala ou luto, este em decorrência de falecimento de pai, mãe, filho, cônjuge, companheira (o) e dependente juridicamente reconhecido.
<br><b>Parágrafo único</b> – Não serão descontadas, no curso de três dias, as faltas do PROFESSOR por motivo de falecimento de sogra, sogro, neto, neta, irmão ou irmão.

<tr><td class=titulo>38. Congressos, simpósios e equivalentes
<tr><td class=campo style="text-align:justify">Os abonos de falta para comparecimento a congressos e simpósios serão concedidos mediante aceitação por parte da MANTENEDORA, que deverá formalizar por escrito a dispensa do PROFESSOR.
<br><b>Parágrafo único</b> - A participação do PROFESSOR nos eventos descritos no caput não caracterizará atividade extraordinária.

<tr><td class="campov" align="center"><b>Outras disposições sobre jornada

<tr><td class=titulo>39. Janelas
<tr><td class=campo style="text-align:justify">Considera-se janela a aula vaga existente no horário do PROFESSOR entre duas outras aulas ministradas no mesmo turno. O pagamento das janelas é obrigatório, devendo o PROFESSOR permanecer à disposição da MANTENEDORA nesses períodos, ressalvada a aceitação pelo PROFESSOR, através de acordo formalizado entre as partes antes do início das aulas, quando as janelas não serão pagas.
<br><b>Parágrafo único</b> - Ocorrendo a hipótese da ressalva supra e caso o PROFESSOR seja solicitado esporadicamente a ministrar aulas ou a desenvolver qualquer outra atividade inerente ao magistério, no horário de janelas não-pagas, essas atividades serão remuneradas como aulas extras, com adicional de 100% (cem por cento).

<tr><td class="campov" align="center"><b>Férias e licenças
<tr><td class="campov" align="center"><b>Férias coletivas

<tr><td class=titulo>40. Férias
<tr><td class=campo style="text-align:justify">As férias anuais dos PROFESSORES serão coletivas, com duração de trinta dias corridos e gozados em julho de 2015. Qualquer alteração deverá ser aprovada por órgão competente, conforme o estabelecido em Estatuto ou Regimento e deverá constar do calendário escolar, obrigatoriamente divulgado aos PROFESSORES até o início de cada período letivo e enviado ao Sindicato.
<br><b>Parágrafo primeiro</b> – A MANTENEDORA está obrigada a pagar o salário das férias e o abono constitucional de 1/3 (um terço) até quarenta e oito horas antes do início das férias.
<br><b>Parágrafo segundo</b> – As férias não poderão ser iniciadas aos domingos, feriados, dias de compensação do descanso semanal remunerado e nem aos sábados, quando estes não forem dias normais de aula.
<br><b>Parágrafo terceiro</b> – Também terá direito às férias coletivas de trinta dias corridos nos períodos estabelecidos no caput, O PROFESSOR que, além de ministrar aulas, tenha cargo de direção ou exerça outras atividades não docentes na MANTENEDORA.
<br>Caso o exercício da atividade administrativa em concomitância com a função docente impossibilite a concessão de férias nos termos do caput, as férias anuais desse PROFESSOR poderão ser gozadas em dois períodos, um deles obrigatoriamente no mês de julho de cada ano.
<br><b>Parágrafo quarto</b> – Na hipótese da divisão das férias anuais do PROFESSOR nos termos do parágrafo anterior, um dos períodos não poderá ser inferior a 10 (dez) dias, sendo proibido o exercício de qualquer atividade nesses períodos.

<tr><td class="campov" align="center"><b>Licença remunerada

<tr><td class=titulo>41. Recesso escolar
<tr><td class=campo style="text-align:justify">O recesso escolar anual é obrigatório e tem duração de trinta dias corridos, gozados preferencialmente no mês de janeiro de 2016.
<br>Durante o recesso escolar anual que não pode, de maneira alguma, coincidir com o período definido para as férias coletivas do ano respectivo, o PROFESSOR não poderá ser convocado para trabalho algum.
<br><b>Parágrafo primeiro</b> – Na vigência da presente Convenção, as instituições cujos calendários escolares, determinados pelo órgão competente conforme o estabelecido em Estatuto ou Regimento, não observarem o determinado pelo caput para o recesso escolar anual dos PROFESSORES, poderão concedê-lo em um período de no mínimo vinte dias corridos em janeiro de 2016 e em no máximo mais três períodos compostos por dias normais de aula e consecutivos, obrigatoriamente no período compreendido entre março de 2015 e fevereiro de 2016.
<br><b>Parágrafo segundo</b> – No caso de os calendários escolares preverem a divisão do recesso escolar dos PROFESSORES, os períodos definidos na conformidade do parágrafo primeiro não poderão ser iniciados aos domingos, feriados, dias de compensação do descanso semanal remunerado e nem aos sábados, quando esses não forem dias normais de aulas.
<br><b>Parágrafo terceiro</b> – As Instituições cujas atividades não possam ser interrompidas, tais como aquelas desenvolvidas em hospital, clínica, laboratório de análise, escritórios experimentais, pesquisas, dentre outros, ou que ministrem cursos em que sejam utilizadas instalações específicas ou que prestem atendimento à comunidade que não pode ser suspenso, poderão conceder aos PROFESSORES o recesso escolar anual definido no caput de maneira escalonada ao longo de cada ano.
<br><b>Parágrafo quarto</b> – Os calendários escolares que definirão os períodos de recesso escolar dos PROFESSORES serão obrigatoriamente divulgados aos PROFESSORES até o início de cada período letivo e enviados ao Sindicato.

<tr><td class="campov" align="center"><b>Licença não remunerada

<tr><td class=titulo>42. Licença sem remuneração.
<tr><td class=campo style="text-align:justify">O PROFESSOR com mais de cinco anos ininterruptos de serviço na MANTENEDORA terá direito a licenciar-se, sem remuneração, por um período máximo de dois anos, não sendo este período de afastamento computado para contagem de tempo de serviço ou para qualquer outro efeito, inclusive legal.
<br><b>Parágrafo primeiro</b> - A licença ou sua prorrogação deverá ser comunicada por escrito, à MANTENEDORA, com antecedência mínima de noventa dias do período letivo, devendo especificar as datas de início e término do afastamento. A licença só terá início a partir da data expressa no comunicado, mantendo-se, até aí, todas as vantagens contratuais. A intenção de retorno do PROFESSOR à atividade deverá ser comunicada à MANTENEDORA, no mínimo, sessenta dias antes do término do afastamento.
<br><b>Parágrafo segundo</b> - O término do afastamento deverá coincidir com o início do período letivo.
<br><b>Parágrafo terceiro</b> - O PROFESSOR que tenha ou exerça cargo de confiança deverá, junto com o comunicado de licença, solicitar seu desligamento do cargo a partir do início do período de licença.
<br><b>Parágrafo quarto</b> - Considera-se demissionário o PROFESSOR que, ao término do afastamento, não retornar às atividades docentes.
<br><b>Parágrafo quinto</b> - Ocorrendo a dispensa sem justa causa ao término da licença, o PROFESSOR não terá direito à “Garantia Semestral de Salários”, prevista na presente Convenção.

<tr><td class="campov" align="center"><b>Outras disposições sobre férias e licenças

<tr><td class=titulo>43. Licença paternidade
<tr><td class=campo style="text-align:justify">A licença paternidade terá duração de cinco dias.

<tr><td class="campot" align="center"><b>Saúde e segurança do trabalhador
<tr><td class="campov" align="center"><b>Uniforme

<tr><td class=titulo>44. Uniformes
<tr><td class=campo style="text-align:justify">A MANTENEDORA deverá fornecer gratuitamente, no mínimo, dois uniformes por ano, quando o seu uso for exigido.

<tr><td class="campov" align="center"><b>Aceitação de atestados médicos

<tr><td class=titulo>45. Atestados médicos e abono de faltas
<tr><td class=campo style="text-align:justify">A MANTENEDORA está obrigada a abonar as faltas dos PROFESSORES, mediante a apresentação de atestados médicos ou odontológicos.

<tr><td class="campot" align="center"><b>Relações sindicais
<tr><td class="campov" align="center"><b>Acesso do sindicato ao local de trabalho

<tr><td class=titulo>46. Quadro de avisos
<tr><td class=campo style="text-align:justify">A MANTENEDORA deverá colocar, nas salas de professores, quadro de aviso à disposição do Sindicato para fixação de comunicados de interesse da categoria, sendo vedada a divulgação de matéria político-partidária ou ofensiva a quem quer que seja.
<br><b>Parágrafo único</b> – O dirigente sindical terá livre acesso à sala dos professores, no horário de intervalo das aulas, para atualizar o material divulgado no quadro de avisos.

<tr><td class="campov" align="center"><b>Representante sindical

<tr><td class=titulo>47. Delegado representante
<tr><td class=campo style="text-align:justify">A MANTENEDORA assegurará a eleição de 1 (um) Delegado Representante para cada Instituição de Ensino Superior mantida, com mandato de 1 (um) ano, que terá a garantia de emprego e salários a partir da inscrição de sua candidatura até o término do semestre letivo em que sua gestão se encerrar.
<br><b>Parágrafo primeiro</b> – A eleição dos Delegados Representantes será realizada pelo Sindicato em cada campus da Instituição de Ensino Superior mantida, por voto direto e secreto. É exigido quórum de 50% (cinquenta por cento) mais um do corpo docente da unidade onde a eleição ocorrer.
<br><b>Parágrafo segundo</b> – O Sindicato comunicará a eleição à MANTENEDORA, com a relação dos candidatos inscritos, com antecedência mínima de sete dias corridos da data da eleição. Nenhum candidato poderá ser demitido a partir da data da comunicação até o término da apuração.
<br><b>Parágrafo terceiro</b> – É condição necessária que os candidatos sejam filiados ao Sindicato e que tenham, à data da eleição, pelo menos um ano de serviço na MANTENEDORA.

<tr><td class="campov" align="center"><b>Liberação de empregados para atividades sindicais

<tr><td class=titulo>48. Assembleias sindicais
<tr><td class=campo style="text-align:justify">Todo PROFESSOR terá direito a abono de faltas para o comparecimento a assembleias da categoria.
<br><b>Parágrafo primeiro</b> – Na vigência desta Convenção, os abonos estão limitados a dois sábados e mais dois dias úteis para cada período compreendido entre o mês de março e o mês de fevereiro do ano subsequente. As duas assembleias realizadas durante os dias úteis deverão ocorrer em períodos distintos.
<br><b>Parágrafo segundo</b> – O Sindicato ou a Federação deverá informar ao SEMESP ou à MANTENEDORA, por escrito, com antecedência mínima de quinze dias corridos. Na comunicação deverão constar a data e o horário da assembleia.
<br><b>Parágrafo terceiro</b> – Os dirigentes sindicais não estão sujeitos ao limite previsto no parágrafo primeiro desta cláusula. As ausências decorrentes do comparecimento às assembleias de suas entidades serão abonadas mediante prévia comunicação formal à MANTENEDORA.
<br><b>Parágrafo quarto</b> – A MANTENEDORA poderá exigir dos PROFESSORES e do dirigente sindical atestado emitido pelo Sindicato ou pela Federação que comprove o seu comparecimento à assembleia.

<tr><td class=titulo>49. Congresso do Sindicato
<tr><td class=campo style="text-align:justify">Na vigência desta Convenção, o Sindicato promoverá um evento de natureza política ou pedagógica (congresso ou jornada). A MANTENEDORA abonará as ausências de seus PROFESSORES que participarem do evento, nos seguintes limites:
<blockquote style="margin-top:0;margin-bottom:0">
a) na unidade de ensino que tenha até 49 (quarenta e nove) PROFESSORES será garantido o abono a um PROFESSOR;
<br>b) na unidade de ensino que tenha entre 50 (cinquenta) e 99 (noventa e nove) PROFESSORES será garantido o abono a 2 (dois) PROFESSORES;
<br>c) na unidade de ensino que tenha mais de 100 (cem) PROFESSORES será garantido o abono a 3 (três) PROFESSORES.
</blockquote>
Tais faltas, limitadas ao máximo em dois dias úteis além do sábado, em cada evento, serão abonadas mediante a apresentação de atestado de comparecimento fornecido pelo Sindicato. O PROFESSOR deverá repor as aulas que, por ventura, sejam necessárias para complementação das horas letivas mínimas exigidas pela legislação.

<tr><td class="campov" align="center"><b>Outras disposições sobre relação entre sindicato e empresa

<tr><td class=titulo>50. Relação nominal
<tr><td class=campo style="text-align:justify">Na vigência desta Convenção, obriga-se a MANTENEDORA a encaminhar ao Sindicato, até o final do mês de junho de cada ano, a relação nominal dos PROFESSORES que integram seu quadro de funcionários, acompanhada do valor do salário mensal e das guias das contribuições sindical e assistencial. A relação poderá ser enviada por meio magnético ou pela internet, ou poderá ainda ser encaminhada cópia da folha de pagamento do mês relativo ao desconto da contribuição sindical.

<tr><td class=titulo>51. Acordos internos - cláusulas mais favoráveis
<tr><td class=campo style="text-align:justify">Ficam assegurados os direitos mais favoráveis decorrentes de acordos internos ou de acordos coletivos de trabalho celebrados entre a MANTENEDORA e o Sindicato.

<tr><td class="campot" align="center"><b>Disposições gerais
<tr><td class="campov" align="center"><b>Regras para a negociação

<tr><td class=titulo>52. Comissão Permanente de Negociação
<tr><td class=campo style="text-align:justify">Fica mantida a Comissão Permanente de Negociação constituída de forma paritária, por três representantes das entidades sindicais (profissional e econômica), com o objetivo de:
<blockquote style="margin-top:0;margin-bottom:0">
a) fiscalizar o cumprimento das cláusulas vigentes;
<br>b) elucidar eventuais divergências de interpretação das cláusulas desta Convenção;
<br>c) discutir questões não contempladas na presente Convenção.
<br>d) deliberar no prazo máximo de trinta dias a contar da data da solicitação protocolizada no SEMESP, sobre modificação de pagamento da assistência médico-hospitalar, conforme os parágrafos 1º e 3º da cláusula “Assistência Médico Hospitalar” desta Convenção e sobre o valor da remuneração da hora-aula, conforme o parágrafo 2º da cláusula “Duração da hora-aula” desta Convenção.
<br>e) criar subsídios para a Comissão de Tratativas Salariais, através da elaboração de documentos, para a definição das funções/atividades e o regime de trabalho dos PROFESSORES.
</blockquote>
	<b>Parágrafo primeiro</b> - As entidades sindicais componentes da Comissão Permanente de Negociação indicarão seus representantes, no prazo máximo de trinta dias corridos, a contar da assinatura desta Convenção.
<br><b>Parágrafo segundo</b> - A Comissão Permanente de Negociação deverá reunir-se mensalmente, no décimo dia útil, às 15 (quinze) horas, alternadamente nas sedes das entidades sindicais que a compõem. No caso específico do item “d“ do caput, deverá haver convocação específica feita pelo SEMESP.

<tr><td class="campov" align="center"><b>Mecanismos de solução de conflitos

<tr><td class=titulo>53. Foro Conciliatório para Solução de Conflitos Coletivos
<tr><td class=campo style="text-align:justify">Fica mantida a existência do Foro Conciliatório que tem como objetivo procurar resolver questões referentes ao não cumprimento de normas estabelecidas na presente Convenção e eventuais divergências trabalhistas existentes entre a MANTENEDORA e seus PROFESSORES.
<br><b>Parágrafo primeiro</b> - O Foro será composto por membros do SEMESP e do Sindicato. As reuniões deverão contar, também, com as partes em conflito que, se assim o desejarem, poderão delegar representantes para substituí-las e/ou serem assistidas por advogados.
<br><b>Parágrafo segundo</b> - O SEMESP e o Sindicato deverão indicar os seus representantes no Foro num prazo de trinta dias a contar da assinatura desta Convenção.
<br><b>Parágrafo terceiro</b> - Cada seção do Foro será realizada no prazo máximo de quinze dias a contar da solicitação formal e obrigatória de qualquer uma das entidades que o compõem, devendo constar na solicitação a data, o local e o horário em que a mesma deverá se realizar. O não comparecimento de qualquer uma das partes acarretará no encerramento imediato das negociações.
<br><b>Parágrafo quarto</b> - Nenhuma das partes envolvidas ingressará com ação na Justiça do Trabalho durante as negociações de entendimento.
<br><b>Parágrafo quinto</b> - Na ausência de solução do conflito ou na hipótese de não-comparecimento de qualquer uma das partes, a comissão responsável pelo Foro fornecerá certidão atestando o encerramento da negociação.
<br><b>Parágrafo sexto</b> - Na hipótese de sucesso das negociações, a critério do Foro, a MANTENEDORA ficará desobrigada de arcar com a multa definida na cláusula “Multa por descumprimento da Convenção”.
<br><b>Parágrafo sétimo</b> - As decisões do Foro terão eficácia legal entre as partes acordantes. O descumprimento das decisões assumidas gerará multa a ser estabelecida no Foro, independentemente daquelas já estabelecidas nesta Convenção.
<br><b>Parágrafo oitavo</b> – Na hipótese de incapacidade econômico-financeira das MANTENEDORAS, os casos serão remetidos para análise e deliberação deste foro.

<tr><td class="campov" align="center"><b>Descumprimento do instrumento coletivo

<tr><td class=titulo>54. Multa por descumprimento da Convenção
<tr><td class=campo style="text-align:justify">O descumprimento desta Convenção obrigará a MANTENEDORA ao pagamento de multa correspondente a 1% (um por cento) do salário do PROFESSOR, para cada uma das cláusulas não cumpridas, acrescidas de juros, a cada PROFESSOR prejudicado.
<br><b>Parágrafo único</b> – A MANTENEDORA está desobrigada de arcar com a multa prevista no caput, caso a cláusula descumprida já estabeleça uma multa pelo seu não cumprimento.

<tr><td class=titulo>55. Contribuição Assistencial
<tr><td class=campo style="text-align:justify">Obriga-se a MANTENEDORA a promover o desconto da contribuição assistencial, na folha de pagamento de seus PROFESSORES, sindicalizados e/ou filiados ou não, para recolhimento em favor do Sindicato profissional, conforme base territorial definida no MTE, em conta especial, na importância deliberada pelas respectivas Assembleias Gerais, se observados os parágrafos abaixo.
<br><b>Parágrafo primeiro</b> – Fica assegurado ao PROFESSOR o direito de oposição à cobrança da contribuição assistencial, a ser exercido, sem qualquer vício de vontade, em 30 (trinta) dias após a entrada em vigor da presente Convenção Coletiva, com o depósito perante o Ministério do Trabalho e Emprego, a ser exercido de modo individual, pessoalmente ou por meio de carta registrada encaminhada ao Sindicato profissional, com cópia à entidade Mantenedora.
<br><b>Parágrafo segundo</b> – O recolhimento da contribuição assistencial será realizado obrigatoriamente pela própria MANTENEDORA, até o 10º (décimo) dia dos meses subsequentes aos descontos, em guias próprias, fornecidas pelo Sindicato da categoria profissional.
<br><b>Parágrafo terceiro</b> - Os Sindicatos representantes das categorias patronal e profissional ficam obrigados a informar, em até 5 (cinco) dias úteis imediatamente após assinatura da Convenção Coletiva, a cada categoria representada (através de publicação em site da entidade na internet, publicação de edital em jornal de ampla circulação na localidade, no quadro de avisos dos empregados na instituição e outros meios eficazes), informações sobre a cobrança da contribuição assistencial e as condições para o exercício de oposição.
<br><b>Parágrafo quarto</b> - A Assembleia para autorização da contribuição assistencial deverá atender aos seguintes requisitos: 1) o edital de convocação da Assembleia Geral deverá ter ampla divulgação, com a publicação em jornais de grande circulação, especialmente convocada para a aprovação da contribuição assistencial, garantindo-se o acesso a todos os trabalhadores, sócios e não sócios; 2) realização em local e horário que facilitem a presença dos trabalhadores; 3) observação dos princípios da proporcionalidade e razoabilidade, para fixação do valor da contribuição assistencial, sendo considerado razoável no ano de 2015, o valor da contribuição correspondente até 1% (um por cento) ao mês, não cumulativa, até 5% (cinco por cento), calculada sobre o valor do salário bruto reajustado por ocasião de cada norma coletiva da categoria.
<br><b>Parágrafo quinto</b> – Para que a contribuição assistencial possa ser pleiteada pelo Sindicato da categoria profissional, o SEMESP deverá receber o edital de convocação e a ata que deliberou sobre a referida contribuição, no prazo de 5 (cinco dias) úteis após a sua realização e anteriormente a inclusão da presente norma no Sistema Mediador.
<br><b>Parágrafo sexto</b> – As Federações representativas dos Sindicatos Profissionais deverão encaminhar ao SEMESP, antes da assinatura da Convenção Coletiva, cópia de eventuais termos de ajustamento de conduta assinados com o Ministério Público ou decisões judiciais acerca de contribuição assistencial.
<br><b>Parágrafo sétimo</b> - O descumprimento de qualquer dos parágrafos anteriores acarretará multa diária de R$ 1.000,00, nos termos do art. 461, 4º do Código de Processo Civil até comprovação de regularização da conduta, sendo revertidos os valores ao FAT – Fundo de Amparo ao Trabalhador.
<br><b>Parágrafo oitavo</b> – Fica expressamente ressalvado que a presente cláusula não prejudica e nem beneficia terceiros que possuam ação judicial ou termo de ajustamento de conduta com entendimento diverso do acima estabelecido, nem a defesa dos direitos individuais de cada trabalhador que se sentir prejudicado.

<tr><td class=campo style="text-align:justify">E por estarem justos e acertados, assinam a presente Convenção Coletiva de Trabalho, a qual será depositada na Delegacia Regional do Trabalho de São Paulo, nos termos do artigo 614 e parágrafos, para fins de arquivo, de modo a surtir, de imediato, os seus efeitos legais.
<br>
<br>São Paulo, 30 de maio de 2015
<br>________________________
<br>Hermes Ferreira Figueiredo
<br>Presidente do Semesp
<br>
<br>


<tr><td class="campot" align="center"><b>ANEXO I
<tr><td class="campov" align="center"><b>REGULAMENTO DO PROGRAMA DE CAPACITAÇÃO
<tr><td class=campo style="text-align:justify">Procedimentos, normas e disposições complementares que regem a concessão, pelo CEBRADE, de gratuidade integral aos filhos ou dependentes legais do PROFESSOR/AUXILIAR, aqui denominados dependentes beneficiários, nos cursos das Instituições de Ensino Superior mantidas e administradas pela MANTENEDORA, na qual o PROFESSOR/AUXILIAR trabalha:
<br><b>1.</b> A instituição que queira aderir ao Termo de Convênio PAET de Concessão de Bolsas de Estudos (ANEXO III) deverá encaminhar ao CEBRADE, o Requerimento de Adesão ao Termo de Convênio (ANEXO II), com pedidos de gratuidade aos dependentes beneficiários nos cursos das Instituições de Ensino Superior (IES) mantidas e administradas pela MANTENEDORA empregadora do PROFESSOR/AUXILIAR, juntamente com o Termo de Convênio PAET (ANEXO III), preenchidos e assinados eletronicamente, para o seguinte endereço eletrônico: convenio.cebrade@semesp.org.br.
<br><b>2.</b> Após o recebimento do Requerimento de Adesão com a indicação dos bolsistas e do Termo de Convênio PAET, preenchidos e assinados pela MANTENEDORA, o CEBRADE fará análise da documentação e, cumpridos os requisitos, enviará a MANTENEDORA, em resposta ao e-mail recebido, cópia do referido termo assinado eletronicamente.
<br><b>3.</b> Sempre que houver ingresso de novos bolsistas, a instituição deverá preencher Termo Aditivo (ANEXO IV) e enviar ao CEBRADE, no mesmo endereço eletrônico mencionado no item I, para que os bolsistas sejam incluídos no Termo de Convênio PAET.
<br><b>4.</b> Caso seja necessário, o CEBRADE, com a supervisão do SEMESP e da FEDERAÇÃO, solicitará ao PROFESSOR/AUXILIAR o envio de documentação que comprove a condição do dependente beneficiário, conforme as condições estabelecidas no item “Programa de capacitação para filhos ou dependentes legais” da cláusula “Bolsas de Estudo” da CCT.
<br><b>5.</b> As gratuidades integrais serão mantidas aos dependentes beneficiários quando o PROFESSOR/AUXILIAR estiver licenciado para tratamento de saúde ou mediante anuência da MANTENEDORA, excetuado o disposto na cláusula “Licença sem Remuneração” da CCT.
<br><b>6.</b> No caso de falecimento do PROFESSOR/AUXILIAR, os dependentes beneficiários continuarão a usufruir as gratuidades integrais até o final do curso, arcando tão somente com as disciplinas cursadas em regime de dependência.
<br><b>7.</b> No caso de dispensa sem justa causa do PROFESSOR/AUXILIAR, os dependentes beneficiários continuarão a usufruir as gratuidades integrais até o final do período letivo.
<br><b>8.</b> Os dependentes beneficiários que forem reprovados no período letivo perderão o direito à bolsa de estudo, voltando a gozar do benefício quando lograrem aprovação naquele período. As disciplinas cursadas em regime de dependência serão de total responsabilidade dos dependentes beneficiários, que deverão arcar com o seu custo.
<br><b>9.</b> Para usufruir as gratuidades integrais dos dependentes beneficiários, não se poderá exigir do PROFESSOR/AUXILIAR pagamento algum, a qualquer título, nem mesmo condicionar a concessão do benefício à associação, sindicalização ou filiação.
<br><b>10.</b> O SEMESP e a FEDERAÇÃO supervisionarão a gestão do Programa pelo CEBRADE e fiscalizarão a disponibilização das bolsas de estudo pela MANTENEDORA, em número suficiente para o atendimento da concessão das gratuidades integrais do PAET nas IES por ela mantida.


<tr><td class="campot" align="center"><b>ANEXO II
<tr><td class="campov" align="center"><b>REQUERIMENTO DE ADESÃO AO TERMO DE CONVÊNIO
<tr><td class=campo style="text-align:justify">
<br>
<br>Ao:
<br>Centro Brasileiro de Desenvolvimento do Ensino Superior - CEBRADE
<br>
<br>A
<br>Entidade Mantenedora, .................., representada neste ato por seu representante legal Sr. ................., portador do RG n.°- .................... - SSP/... e do CPF n° ............., com sede na ...................., vem, por meio da presente, nos termos do que estabelece a Convenção Coletiva de Trabalho e Regulamento do Programa de Capacitação, requerer a adesão ao Termo de Convênio PAET de Concessão de Bolsas de Estudo, cujos alunos participantes seguem abaixo:
<br>

<tr><td class=campo style="text-align:center">

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=600>
<tr><td class=campo>Nome do aluno</td>
<td class=campo>Matrícula</td>
<td class=campo>Curso</td>
<td class=campo>Série</td>
<td class=campo>Porcentagem de bolsa concedida</td>
</tr>
<%for a=1 to 9%>
<tr><td class=campo>&nbsp;</td><td class=campo></td><td class=campo></td><td class=campo></td><td class=campo></td></tr>
<%next%>
</table>
<br>
<br>
<br>
<br>__________________________________________
<br>(Assinatura do representante legal da Mantenedora)
<br>

<tr><td class="campot" align="center"><b>ANEXO III
<tr><td class="campov" align="center"><b>TERMO DE CONVÊNIO PAET DE CONCESSÃO DE BOLSAS DE ESTUDO
<tr><td class=campo style="text-align:justify">
Pelo presente instrumento, de um lado CENTRO BRASILEIRO DE DESENVOLVIMENTO DO ENSINO SUPERIOR – CEBRADE, pessoa jurídica de direito privado, sem fins lucrativos, inscrita no CNPJ sob n.º .............., domiciliada na Rua Cipriano Barata, 2431 – Ipiranga – São Paulo – SP, representado neste ato pelo ..........................................., doravante denominado CEBRADE e de outro lado a xxxxxxxxxxx, entidade doravante denominada abreviadamente INSTITUIÇÃO, representada neste ato por seu ................. Sr. ................., portador do RG n.°- .................... - SSP/... e do CPF n° ............., com sede na ...................., considerando a necessidade de implementar um sistema de concessão de bolsas aos dependentes de professores e auxiliares da educação superior mediante o desenvolvimento do Programa de Amparo Educativo Temporário – PAET, que priorize o desenvolvimento, integração e acesso à Educação Superior no Estado São Paulo, resolvem celebrar o presente convênio de cooperação, e de acordo com as cláusulas e condições a seguir:
<br>
<br><b>DO OBJETO</b>
<br><b>CLÁUSULA PRIMEIRA</b>
<br>O presente Convênio tem por objeto estabelecer, em regime de cooperação mútua entre os partícipes, o desenvolvimento da educação superior no país mediante a concessão de bolsas de estudo aos dependentes legais dos empregados das instituições de ensino superior participantes do presente convênio.
<br>
<br><b>DAS CONDIÇÕES GERAIS</b>
<br><b>CLÁUSULA SEGUNDA</b>
<br>Fica estabelecido entre as partes que o CEBRADE – Centro Brasileiro de Desenvolvimento do Ensino Superior – que possui como um dos seus objetivos, desenvolvimento do Programa de Amparo Educativo Temporário – PAET, concedendo bolsas de estudo em Instituições Privadas de Ensino Superior concederá aos filhos ou dependentes legais do empregado o direito de usufruir as gratuidades integrais do PAET, sem qualquer ônus, nos cursos de graduação e sequencial existentes e administrados pela INSTITUIÇÃO para a qual o empregado trabalha, observado o disposto neste instrumento.
<br><b>PARÁGRAFO PRIMEIRO</b>. A INSTITUIÇÃO deverá disponibilizar ao CEBRADE, mediante requerimento, bolsas de estudo em número suficiente para o atendimento da concessão das gratuidades integrais do PAET nas Instituições de Ensino Superior por ela mantida, para filhos ou dependentes legais dos seus empregados, observada a limitação estabelecida na cláusula de bolsas de estudo.
<br><b>PARÁGRAFO SEGUNDO</b>. Para a concessão das gratuidades integrais aos filhos e dependentes legais do empregado, o CEBRADE não poderá fazer qualquer outra exigência a não ser o comprovante de aprovação no processo seletivo da INSTITUIÇÃO empregadora e a observância dos preceitos estabelecidos neste instrumento.
<br><b>PARÁGRAFO TERCEIRO</b>. Terão direito a requerer e obter do CEBRADE a concessão de bolsas integrais de estudo, os dependentes legais do empregado reconhecidos pela Legislação do Imposto de Renda, ou que estejam sob a sua guarda judicial e vivam sob sua dependência econômica, devidamente comprovada.
<br><b>PARÁGRAFO QUARTO</b>. Os filhos do empregado terão direito a obter do CEBRADE concessão de bolsas de estudo integrais, desde que, na data de efetivação da matrícula no curso superior, não tenham 25 (vinte e cinco anos) completos ou mais.
<br><b>PARÁGRAFO QUINTO</b>. As bolsas de estudo são válidas para cursos de graduação e sequenciais e a INSTITUIÇÃO está obrigada a conceder, no máximo, duas bolsas de estudo por empregado.
<br><b>PARÁGRAFO SEXTO</b>. O beneficiário bolsista, concluinte de curso de graduação não poderá obter nova concessão de gratuidade na mesma instituição.
<br><b>PARÁGRAFO SÉTIMO</b>. As bolsas de estudo serão mantidas aos dependeste quando o empregado estiver licenciado para tratamento de saúde ou em gozo de licença mediante anuência da INSTITUIÇÃO, excetuado quando o empregado tiver licenciado por “Licença sem Remuneração”.
<br><b>PARÁGRAFO OITAVO</b>. No caso de falecimento do empregado, os dependentes legais que já se encontrarem estudando na INSTITUIÇÃO continuarão a gozar das bolsas de estudo até o final do curso.
<br><b>PARÁGRAFO NONO</b>. No caso de dispensa sem justa causa do empregado durante o período letivo, ficam garantidas até o final do período letivo, as bolsas de estudo já existentes.
<br><b>PARÁGRAFO DÉCIMO</b>. Os bolsistas que forem reprovados no período letivo perderão o direito à bolsa de estudo, voltando a gozar do benefício quando lograrem aprovação no referido período. As disciplinas cursadas em regime de dependência serão de total responsabilidade do bolsista, arcando o mesmo com o seu custo.
<br><b>PARÁGRAFO DÉCIMO PRIMEIRO</b>. Além dos casos previstos nesta cláusula, a INSTITUIÇÃO poderá fornecer outras bolsas de estudos, cujas condições serão objeto de termo aditivo a ser firmado entre a INSTITUIÇÃO e o CEBRADE, nos termos do ANEXO IV.
<br>
<br><b>DA COMISSÃO DE ACOMPANHAMENTO DO CONVÊNIO</b>
<br><b>CLÁUSULA TERCEIRA</b>
<br>O SEMESP e a FEDERAÇÃO fiscalizará o CEBRADE na gestão do Programa de Amparo Educativo Temporário para os filhos e dependentes legais dos empregados nas instituições de ensino pertencentes a sua categoria representativa.
<br><b>PARÁGRAFO ÚNICO.</b> Os convenentes desde já expressam concordância quanto à fiscalização, bem como se comprometem a fornecer todos os documentos que lhe forem solicitados para comprovar o cumprimento das obrigações ora assumidas.
<br>
<br><b>DO PRAZO</b>
<br><b>CLÁUSULA QUARTA</b>
<br>O presente Convênio vigorará até 28 de fevereiro de 2015, tendo como termo inicial a data de sua assinatura, podendo ser renovado no interesse dos partícipes por novos prazos.
<br>
<br><b>DO DESCUMPRIMENTO DAS OBRIGAÇÕES</b>
<br><b>CLÁUSULA QUINTA</b>
<br>O descumprimento pelos convenentes dos compromissos assumidos neste convênio ensejará a rescisão do presente instrumento e a aplicação das penalidades previstas na Lei.
<br>
<br><b>CONFIDENCIALIDADE</b>
<br><b>CLÁUSULA SEXTA</b>
<br>Comprometem-se as partes a proteger as informações confidenciais, no caso do presente instrumento dados pessoais e qualquer outro informado na “Solicitação de bolsa de estudo”, sob pena de responder pelos danos causados, sem prejuízo de indenização e outras medidas cabíveis.
<br>
<br><b>DO FORO</b>
<br><b>CLÁUSULA SÉTIMA</b>
<br>E, por estarem os convenentes certos e acordados quanto às cláusulas e condições deste convênio, firmam o presente termo em 2 (duas) vias de igual teor e para um só efeito na presença das testemunhas abaixo assinadas e qualificadas.
<br>
<br>São Paulo ____ de _______, de 2013.
<br>________________________________
<br>CEBRADE
<br>
<br>_________________________________
<br>MANTENEDORA
<br>TESTEMUNHA 1: ____________________________________
<br>RG:_______________________________________________
<br>CPF: ______________________________________________
<br>TESTEMUNHA 2: ____________________________________
<br>RG:_______________________________________________
<br>CPF: ______________________________________________
<br>
<br>

<tr><td class="campot" align="center"><b>ANEXO IV
<tr><td class="campov" align="center"><b>TERMO ADITIVO DE INCLUSÃO DE ALUNO NO CONVÊNIO PAET DE CONCESSÃO DE BOLSAS DE ESTUDO
<tr><td class=campo style="text-align:justify">
<br>
<br>Ao CEBRADE
<br>
<br>A
<br>Entidade Mantenedora, .................., representada neste ato por seu representante legal Sr. ................., portador do RG n.°- .................... - SSP/... e do CPF n° ............., com sede na ...................., vem, por meio da presente, nos termos do que estabelece a Convenção Coletiva de Trabalho e Regulamento da Cláusula de Bolsa de Estudos, solicitar a inclusão dos alunos abaixo indicados no Termo de Convênio PAET de Concessão de Bolsas de Estudos:

<tr><td class=campo style="text-align:center">

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=600>
<tr><td class=campo>Nome do aluno</td>
<td class=campo>Matrícula</td>
<td class=campo>Curso</td>
<td class=campo>Série</td>
<td class=campo>Porcentagem de bolsa concedida</td>
</tr>
<%for a=1 to 9%>
<tr><td class=campo>&nbsp;</td><td class=campo></td><td class=campo></td><td class=campo></td><td class=campo></td></tr>
<%next%>
</table>
<br>
<br>
<br>_____________________________
<br>(Assinatura do representante legal da Mantenedora)
<br>
<br>São Paulo, ___de ___ de 2013
<br>
<br>

<tr><td class="campot" align="center"><b>ANEXO V
<tr><td class="campov" align="center">
<tr><td class=campo style="text-align:justify">
<br>MINISTÉRIO PÚBLICO DO TRABALHO -
<br>PROCURADORIA REGIONAL DO TRABALHO DA 2ª REGIÃO/SP
<br>EXCELENTÍSSIMO SENHOR DOUTOR DESEMBARGADOR PRESIDENTE DO EGRÉGIO TRIBUNAL REGIONAL DO TRABALHO DA 2ª REGIÃO
<br>PROT. 22828 P49 ACORDÃO 20111091459
<br>Julgado com recurso
<br>Ser RECEPÇÃO PROC. RECURSAL
<br>PROC. 0135900382065020074
<br>
<br>
<br>MINISTÉRIO PÚBLICO DO TRABALHO - PROCURADORIA REGIONAL DO TRABALHO DA 2ª REGIÃO, autor da presente ação e, FEDERAÇÃO DOS TRABALHADORES EM ESTABELECIMENTOS DE ENSINO DO ESTADO DE SÃO PAULO – FETEESP e SINDICATO DAS ENTIDADES MANTENEDORAS DE ESTABELECIMENTOS DE ENSINO SUPERIOR NO ESTADO DE SÃO PAULO – SEMESP, rés no presente feito, nos autos do processo supra, vem presente Vossa Excelência para expor e requerer o seguinte:
<br>
<br><b>1º Nos autos do processo supra fora prolatada decisão de primeira instância da 74ª Vara do Trabalho de São Paulo (de 4/9/2007) determinando aos réus:</b>
<br>... “a se absterem de arrecadar contribuições sindicais, previstas em instrumentos normativos negociais dos trabalhadores não filiados, ressalvada expressa autorização dos mesmos, sob pena de pagamento de multa diária de R$ 1.000,00, nos termos do artigo 461, par. 4º do Código de Processo Civil.”
<br>Condeno, ainda, os requeridos a não estipularem, em instrumentos normativos negociais, cláusulas com o fim de arrecadar contribuições sindicais dos trabalhadores não filiados, ressalvada expressa autorização dos mesmos, sob pena de multa de R$ 50.000,00 por cláusula que vier a ser estipulada nesse sentido.
<br>As multas eventualmente impostas serão revertidas ao Fundo de Amparo ao Trabalhador – FAT ...”
<br>
<br><b>2º Em julgamento de recurso ordinário interposto da daquela Decisão de 1º grau, à 15ª Turma do Egrégio Tribunal Regional do Trabalho, confirmou a sentença em votação Unanime acompanhando o Voto da Relatora designada, podendo ser destacado de tal decisão o seguinte:</b>
<br>“... Sempre entendemos que as contribuições assistenciais, previstas nas Convenções Coletivas de Trabalho, são devidas por todos os empregados representados pelo sindicato autor, independentemente de serem associados à entidade sindical. Isso porque no sistema sindical brasileiro o sindicato representa a totalidade da categoria profissional e não apenas os seus associados, de forma que quando é prolatada sentença normativa, são desses instrumentos beneficiários todos os membros da categoria, independentemente de sua filiação ao sindicato. Para os associados resta o ônus de contribuir com as mensalidades dos
sindicatos, beneficiando-se de sua associação à entidade. Sob a nossa ótica, esses sistema não fere a liberdade sindical, vez que a Constituição Federal, apesar de ter elevado à categoria constitucional o princípio da liberdade sindical, manteve e também elevou a tal categoria, o sistema de unicidade sindical. Assim, cabe a um único sindicato por categoria e base territorial a representação de todos os empregados, independentemente se sua filiação, como visto acima. Consequência desse sistema é autorização para que o ente sindical estabeleça contribuição assistencial, para despesas com negociações coletivas em prol de toda a categoria. Em nosso entendimento, não é o caso de aplicação do precedente 119 do C.TST, dirigido às ações em dissídio coletivo, aqui se tratando de aplicação de cláusula convencional já fixada.
<br>(...)
<br>Conforme cláusulas habitualmente concedidas pelo grupo normativo do TRT 2ª Região, foi editado o Precedente 21, da E. SDC, com a seguinte redação: DESCONTO ASSISTENCIAL – desconto assistencial de 5% dos empregados, associados ou não, de uma só vez e quando do primeiro pagamento dos salários já reajustados, em favor da entidade de trabalhadores, importância essa a ser recolhida em conta vinculada sem limite à Caixa Econômica Federal”.
<br>Verifica-se que a Convenção Coletiva de 2005 observou o limite de 5% estabelecido no Precedente acima citado, que também se refere a empregados associados ou não. Ocorre que Convenção Coletiva de Trabalho, que prevê descontos compulsórios de contribuição assistencial entre trabalhadores, deveria também ter estipulado cláusula conferindo ao trabalhador o exercício do direito de oposição, possibilitando a manifestação de sua discordância em relação aos descontos.
<br>Diante disso, nada a modificar na r. sentença que condicionou os descontos dos trabalhadores não filiados à expressa manifestação dos mesmos, tendo em vista a ausência de cláusula estabelecendo o direito de oposição.” (...);
<Br>3 – As rés, em face do V. Acordão acima mencionado, apresentam embargos de declaração que foram acolhidos parcialmente para:
<br>“(...)
<br>4. Da multa diária e multa por descumprimento.
<Br>Com relação à alegação recursal no sentido de que a ação civil pública não comporta multa (fls. 346/347), há omissão que passa a ser sanada.
<br>A r. sentença condenou as reclamadas a: a) se absterem de arrecadar contribuições, previstas em instrumentos normativos negociais, dos trabalhadores não filiados, ressalvada expressa autorização dos mesmos, sob pena de multa diária do valor de R$ 1.000,00, (art. 461, par. 4º do CPC) e b) não estipularem em instrumentos normativos negociais cláusulas com o fim de arrecadar contribuições sindicais dos trabalhadores não filiados, ressalvada expressa autorização dos mesmos, sob pena de multa de R$ 50.000,00 por cláusula que vier a ser estipulada nesse sentido (fls.244).
<br>A aplicação de multa encontra amparo no art. 21 da Lei nº 7.347/85 (Lei da Ação Civil Pública), que remete ao título III da Lei n. 8.078/90 (Código de Defesa do Consumidor). Este ultimo trata de aspectos processuais, dispondo, em seu art. 84, a respeito da tutela específica, prevendo, inclusive, a aplicação da multa.
<Br>Não se justifica, também, a diminuição do valor arbitrado na origem, tendo em vista que a aplicação das multas não se destina a fazer com que o devedor as pague, mas sim forçar o cumprimento da obrigação na forma específica” (...).
<br>4 – Atualmente, a decisão proferida no V. Acórdão que julgou o recurso ordinário e confirmada no julgamento dos embargos declaratórios opostos pelas rés, não transitou em julgado e o feito encontra-se pendente de análise de admissibilidade do recurso de revista interposto pelas demandadas:
<br>5 – destarte, considerando os riscos do processo, outrossim, diante dos termos da R. Sentença recorrida e do entendimento consignado no V. Acórdão acima citado, que acrescentou fundamentação nova à Decisão de 1º grau, sem alterar entretanto o decisum, os signatários vêm à presença do V. Excelência, para informar que se compuseram para por fim à demanda, sendo que as rés, para adequação dos futuras normas coletivas a serem produzidas ao entendimento da jurisprudência dominante desta Corte, incluindo o pensamento exposto no V. Acórdão acima citado e consubstanciado também no Precedente Normativo n.21 do Tribunal Regional do Trabalho da 2ª Região, se comprometem a:
<blockquote style="margin-top:0;margin-bottom:0">
a) se absterem de estipular em instrumentos contratuais coletivos de trabalho, incluindo-se também aqueles instrumentos firmados em nome dos sindicatos filiados à federação profissional signatária, e/ou com anuência desta, cláusulas prevendo contribuições por participação em negociações coletivas (negocial/assistencial) dos trabalhadores não filiados a entidade sindical sem garantir o exercício do direito de oposição a cobrança de tais contribuições, sob pena de pagamento de multa diária do valor de R$ 1.000,00, nos termos do art. 461, 4º do Código de Processo Civil até comprovação de regularização da conduta, sendo revertidos os valores ao FAT – Fundo de Amparo ao Trabalhador;
<br>b) que a instituição de contribuição assistencial/negocial em cada norma contratual coletiva será aprovada em assembleia geral da categoria convocada para este fim, com ampla divulgação, garantida a participação de sócios e não sócios, realizada em local e horário que facilitem a presença dos trabalhadores, sendo que as rés observarão os princípios da proporcionalidade e razoabilidade, para fixação do valor da contribuição assistencial, sendo que para efeitos do presente acordo, é considerado razoável o valor da contribuição correspondente até 1% (um por cento) ao mês, não cumulativa, até 5% (cinco por cento) por ano de vigência da norma contratual coletiva, calculada sobre o valor do salário bruto reajustado por ocasião de cada norma coletiva da categoria, sob pena de pagamento de multa diária de R$ 1.000,00, nos termos do art. 461, 4º do Código de Processo Civil até comprovação de regularização da conduta, sendo revertidos os valores ao FAT – Fundo de Amparo ao Trabalhador;
<br>c) as rés assegurarão, ao trabalhador integrante da categoria o direito de oposição à cobrança da contribuição assistencial/negocial fixada em cada norma contratual coletiva, a ser exercido, sem qualquer vício de vontade, em prazo razoável, que para efeitos tão somente do presente acordo fica estabelecido em 30 (trinta) dias após a entrada em vigor da norma contratual coletiva com o depósito perante o Ministério do Trabalho e Emprego (acordo/convenção coletiva de trabalho) a ser exercido de modo individual, pessoalmente ou por meio de carta encaminhada à entidade profissional ré, com cópia à entidade Mantenedora, sob pena de pagamento de multa diária de R$ 1.000,00 nos termos do artigo 461, 4º do Código de Processo Civil ate a comprovação de regularização da conduta, sendo revertidos os valores ao FAT – Fundo de Amparo ao Trabalhador;
<br>d) para efeito da cobrança da contribuição assistencial/negocial as rés se comprometem, em 5 (cinco) dias úteis, imediatamente após a pactuação do instrumento coletivo de trabalho, a divulgar a celebração do acordo ou convenção coletiva e trabalho perante a categoria respectivamente representada (através de publicação em site da entidade na internet, publicação de edital em jornal de ampla circulação na localidade e outros meios eficazes) , incluindo informações sobre a cobrança das referidas contribuições e para condições de exercício de oposição, sob pena de pagamento de multa diária de R$ 1.000,00, nos termos do artigo 461, 4º do Código de Processo Civil até a comprovação de regularização da conduta, sendo revertidos os valores ao FAT – Fundo de Amparo ao Trabalhador;
<br>e) para efeito da contribuição assistencial prevista em instrumento coletivo de trabalho, o SEMESP deverá receber o edital de convocação e a ata que deliberou sobre a referida contribuição, no prazo de 5 (cinco dias) úteis após a sua realização. O edital de convocação deverá ser publicado em jornais de grande circulação, garantindo-se o acesso a todos os trabalhadores;
<br>f) as federações representativas de sindicatos profissionais deverão encaminhar ao SEMESP, antes de qualquer assinatura de convenção coletiva, cópias de termos de ajustamento de conduta assinados com o Ministério Público ou decisões judiciais acerca de contribuição assistencial, sob pena de multa diária de R$ 1.000,00, sendo revertidos os valores ao FAT – Fundo de Amparo ao Trabalhador;
<br>g) indenização no valor de R$ 50.000,00, a titulo de reparação do dano moral coletivo, por cláusula que vier a ser confeccionada em cada instrumento contratual coletivo, contrariando e estipulado nas letras “a” a “d” supra, sendo revertidos os valores ao FAT – Fundo de Amparo ao Trabalhador;
<br>h) fica expressamente ressalvado que o presente acordo não prejudica e nem beneficia terceiros que possuam ação judicial ou termo de ajustamento de conduta com entendimento diverso do acima estabelecido, nem a defesa dos direitos individuais de cada trabalhador que se sentir prejudicado;
<br>i) custas e demais despesas processuais ficam à cargo das rés;
</blockquote>
5 – destarte requerem a homologação do presente acordo para que produza os seus devidos efeitos legais, desistindo as rés do recurso de revista interposto.



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