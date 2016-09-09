/******************************************
Menu item creation:
myCoolMenu.makeMenu(name, parent_name, text, link, target, width, height, regImage, overImage, regClass, overClass , align, rows, nolink, onclick, onmouseover, onmouseout) 
*************************************/

oCMenu.makeMenu('top3','','RH Benefícios','','',90)
	oCMenu.makeMenu('sub31','top3','Bolsas de Estudo','','',120)
		oCMenu.makeMenu('sub312','sub31','Listagem geral',     'bolsa/rh_bolsalista.asp',   '')
		oCMenu.makeMenu('sub313','sub31','Agrupados por tipo', 'bolsa/rh_bolsagrupo.asp',   '')
		oCMenu.makeMenu('sub314','sub31','Lista por bolsista', 'bolsa/rh_bolsistas.asp',    '')
		oCMenu.makeMenu('sub315','sub31','Resumo',             'bolsa/rh_bolsaresmo.asp',   '')
		oCMenu.makeMenu('sub316','sub31','Resumo p/ curso',    'bolsa/rh_bolsarescurso.asp','')
		oCMenu.makeMenu('sub317','sub31','Controle Todos',     'bolsa/rhbolsafunc_c.asp','')

	oCMenu.makeMenu('sub32','top3','Convênios','','',120)
		oCMenu.makeMenu('sub325','sub32','Conveniados Fieo',         'convenio/rhconveniados.asp',    '',150)
		oCMenu.makeMenu('sub326','sub32','Conveniados Faculdades',   'convenio/rhconveniadosfac.asp', '',150)
		oCMenu.makeMenu('sub327','sub32','Conveniados Fac.p/curso',  'convenio/rhconveniadosfacc.asp','',150)
		oCMenu.makeMenu('sub328','sub32','Conveniados Fac.aprovados','convenio/rhconveniadosfacf.asp','',150)

	oCMenu.makeMenu('sub33','top3','Ticket Car','','',120)
		oCMenu.makeMenu('sub331','sub33','Recibo de cartão','ticketcar/rhform_recibotcar.asp',  '')
		oCMenu.makeMenu('sub332','sub33','Controle Pedido', 'ticketcar/tc_controle.asp',  '')

	oCMenu.makeMenu('sub34','top3','V.Transporte','','',120)
		oCMenu.makeMenu('sub344','sub34','Controle de Saldo','vt/vt_saldo.asp',  '',130)
			oCMenu.makeMenu('sub3433','sub344','Razão Analítico','vt/vt_razao.asp',  '',130)
			oCMenu.makeMenu('sub3434','sub344','Saldo','vt/vt_saldo.asp',  '',130)
		oCMenu.makeMenu('sub345','sub34','Op.SPTRANS',       'vt/sptrans_1.asp', '',130)

oCMenu.makeMenu('top4','','RH Controles','','',90)
	oCMenu.makeMenu('sub40','top4','Seleção','','')
		oCMenu.makeMenu('sub401','sub40','Alteração de Função',  'cs/rhaltfuncao.asp',         '',130)
		oCMenu.makeMenu('sub402','sub40','Requisição de Pessoal','cs/rhrequisicao.asp',        '',130)

	oCMenu.makeMenu('sub41','top4','INSS','','')
		oCMenu.makeMenu('sub413','sub41','Carta ao Professor', 'inss/controleteto_pedido.asp', '')
		oCMenu.makeMenu('sub415','sub41','Carta Teto Mod.2',   'inss/cartateto2.asp',         '')
		oCMenu.makeMenu('sub417','sub41','Carta Restituição',  'inss/cartarest.asp',          '')
		oCMenu.makeMenu('sub418','sub41','Carta Rest. Aut.',   'inss/cartarestaut.asp',  '')
		oCMenu.makeMenu('sub419','sub41','Solicitações',       'inss/pedidoteto.asp',          '')

	oCMenu.makeMenu('sub42','top4','Formulários','','')
		oCMenu.makeMenu('sub422','sub42','Desconto Estac.',   'forms/rhform_autveiculo.asp',    '')

	oCMenu.makeMenu('sub43','top4','C.& Salários','','')
		oCMenu.makeMenu('sub431','sub43','Descrição de Cargos',   'cs/rhdescricaocs.asp',  '',120)
		oCMenu.makeMenu('sub432','sub43','Descrição de Cargos X', 'cs/rhdescricaocs2.asp', '',120)
		oCMenu.makeMenu('sub436','sub43','Conferência CS-Prof.', 'cs/conf_cs_prof.asp', '',120)

	oCMenu.makeMenu('sub44','top4','Listagens','','')
		oCMenu.makeMenu('sub441','sub44','Conf. Horário',          'diversos/rhhorarios.asp',    '',130)

	oCMenu.makeMenu('sub45','top4','Professores','','')
		oCMenu.makeMenu('sub451','sub45','Cartão Professor',          'professor/cartao.asp',         '',130)
		oCMenu.makeMenu('sub454','sub45','Relatórios','','',130)
			oCMenu.makeMenu('sub4541','sub454','por chapa',       'professor/rpt_chapa.asp','',130)
			oCMenu.makeMenu('sub4542','sub454','por curso',       'professor/rpt_curso.asp','',130)
			oCMenu.makeMenu('sub4543','sub454','Capa para pontos','professor/rpt_capa.asp', '',130)
			oCMenu.makeMenu('sub4544','sub454','da Pós',          'professor/rpt_pos.asp',  '',130)
		oCMenu.makeMenu('sub458','sub45','Teste',                     'professor/teste.asp','',130)
		oCMenu.makeMenu('sub457','sub45','Cadastro Prof.Vis.',        'professor/cadprofessor.asp','',130)
		oCMenu.makeMenu('sub459','sub45','Cartão Professor Apont',    'professor/cartao_ap.asp',         '',130)
		oCMenu.makeMenu('sub460','sub45','Cartão Professor Base teste','professor/cartao2.asp',         '',130)
		oCMenu.makeMenu('sub460a','sub45','Aviso Prévio',            'professor/avisoprevio.asp','',130)

	oCMenu.makeMenu('sub46','top4','Autônomos','','')
		oCMenu.makeMenu('sub462','sub46','Impressão RPA', 'autonomo/imprime_rpa.asp',   '',130)
		oCMenu.makeMenu('sub464','sub46','Impressão DARF','autonomo/imprime_darf.asp',  '',130)

oCMenu.makeMenu('top7','','Supervisor','','',70)
	oCMenu.makeMenu('sub71','top7','Usuários',    'users.asp',       '',120)

oCMenu.makeMenu('top8','','Secretaria de Curso','','',130)
	oCMenu.makeMenu('sub83','top8','Checagens','','','120')
		oCMenu.makeMenu('sub832','sub83','Grade Curricular','grades/gradecurricular.asp','','120')
		oCMenu.makeMenu('sub833','sub83','Histórico Disciplinas','grades/tabdisciplina.asp','','120')
		oCMenu.makeMenu('sub834','sub83','C.H.-Diminuição',      'grades/tabdiminuicao.asp','','120')
		oCMenu.makeMenu('sub835','sub83','C.H.-Aumento',      'grades/tabaumento.asp','','120')
		oCMenu.makeMenu('sub836','sub83','C.H.-Analise por curso',      'grades/tabmesma.asp','','120')

	oCMenu.makeMenu('sub861','top8','Dias Letivos'     , 'secretaria/diasletivos.asp' ,'',120)
	oCMenu.makeMenu('sub871','top8','Dependências'     , 'secretaria/dependencia.asp','',120)
	oCMenu.makeMenu('sub881','top8','Outras Atividades', 'secretaria/atividades.asp'  ,'',120)

oCMenu.makeMenu('top9','','Secretaria Pós','','',130)
	oCMenu.makeMenu('sub90','top9','Grades',            'gradespos/grades.asp' ,'',120)
	oCMenu.makeMenu('sub94','top9','Livro de Ponto','gradespos/ponto.asp','','120')
	oCMenu.makeMenu('sub941','top9','Folha de Ponto','gradespos/pontob.asp','','120')


oCMenu.makeMenu('top10','','','forms/funcionarios.asp','',15,15,'images/personal.gif')

oCMenu.makeMenu('top11','','Imprimir o quadro','','',15,15,'images/printer.gif','','','' ,'','','','myprint()') 

oCMenu.makeMenu('top12','','','mapasite.asp','',15,15,'images/earth.gif')
oCMenu.makeMenu('top13','','','estacionamento/cadveiculo.asp','',15,15,'images/truck.gif')


top.frmMain.location.reload() 