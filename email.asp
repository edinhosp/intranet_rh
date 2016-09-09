<%@ Language=VBScript %>
<!-- #Include file="ADOVBS.INC" -->
<!-- #Include file="funcoesclear.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<title>Envio de email</title>
<body>
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>

<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
set rs.ActiveConnection = conexao

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

sql1="select top 1 f.chapa, f.nome, pnome=intranet_rh.dbo.primeironome(f.nome), p.sexo, p.email " & _
"from corporerm.dbo.PFUNC f inner join corporerm.dbo.PPESSOA p on p.CODIGO=f.CODPESSOA " & _
"and f.CODSECAO='03.1.009' and CODSITUACAO<>'D' and EMAIL is not null and CHAPA in ('02379','00259') "

sql1="select top 10 mataluno, nome, pnome, sexo, email from pesquisa where mataluno>'00000001' order by mataluno"

rs.Open sql1, ,adOpenStatic, adLockReadOnly
'*************** inicio teste **********************
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>" & lcase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext:loop
response.write "</table>"
response.write "# " & rs.recordcount & "<br>"
rs.movefirst
'*************** fim teste **********************

if rs.recordcount>0 then 
do while not rs.eof
	if rs("sexo")="M" then t1="o" else t1="a"
	email=rs("email")
	cabecalho="<html><style type='text/css'>" & _
	"<!--" & _
	"td.titulo { font-size:12pt; font-family:tahoma; font-weight:bold; background-color:Silver; color:Black;} "& _
	"td.campo { font-size:10pt; font-family:tahoma; font-weight:normal; background-color:White; font-size-adjust:inherit; font-stretch:inherit;} " & _
	"p { font-size:10pt; font-family:tahoma; font-weight:normal;} " & _
	"-->"&_
	"</style><body>"
	intro="<table border='1' bordercolor='#a9a9a9' cellpadding='5' cellspacing='0' style='border-collapse: collapse' width=600'>" & _
	"<tr><td class='titulo'>Pesquisa sobre conclusão de curso tecnológico - UNIFIEO</td></tr>" & _
	"<tr><td class='campo'>" & _
	"<p style='margin-bottom:0;margin-top:15'>Prezad" & t1 & " " & rs("pnome") & "<br>" & _
	"<p style='margin-bottom:0;margin-top:15'>Estamos realizando um estudo acadêmico sobre os alunos egressos dos cursos tecnológicos do UNIFIEO e gostaríamos de conta com a sua participação.<br>" & _
	"<p style='margin-bottom:0;margin-top:15'>Levará apenas 2 minutos para responder e a sua participação será indispensável para a mensuração do estudo.<br>" & _
	"<p style='margin-bottom:0;margin-top:15'>Agradecemos a sua colaboração.<br>" & _
	"</td></tr>"
	texto="<tr><td class='campo'>Se tiver problemas para visualizar este formulário, você poderá preenchê-lo on-line:<br>" & _
	"<a href='https://docs.google.com/forms/d/1pALiC_5H732umOqohpIewFQAv_FcThQrjzU8jS-Elo0/viewform'>https://docs.google.com/forms/d/1pALiC_5H732umOqohpIewFQAv_FcThQrjzU8jS-Elo0/viewform</a> " & _
	"" & _
	"</td></tr> <tr><td>"
	'response.write cabecalho & intro & texto

bloco1="<div class='form-body' style=''><div class='ss-form' style=''><form action='https://docs.google.com/forms/d/1pALiC_5H732umOqohpIewFQAv_FcThQrjzU8jS-Elo0/formResponse' method='POST' id='ss-form' target='_self' style=''>" & _
"<div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-select' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_1736008929' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>1) Curso que concluiu?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<select name='entry.1736008929' id='entry_1736008929' style=''><option value='' style=''></option><option value='Tec. Gestão de Recursos Humanos' style=''>Tec. Gestão de Recursos Humanos</option><option value='Tec. Gestão Financeira' style=''>Tec. Gestão Financeira</option><option value='Tec. Gestão Comercial' style=''>Tec. Gestão Comercial</option><option value='Tec. Gestão de Logística' style=''>Tec. Gestão de Logística</option><option value='Tec. Gestão de Eventos' style=''>Tec. Gestão de Eventos</option><option value='Tec. Gestão de Marketing' style=''>Tec. Gestão de Marketing</option><option value='Tec. Secretariado' style=''>Tec. Secretariado</option><option value='Outro' style=''>Outro</option></select></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-select' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_816575205' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>Ano de Conclusão" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<select name='entry.816575205' id='entry_816575205' style=''><option value='' style=''></option><option value='2012' style=''>2012</option><option value='2011' style=''>2011</option><option value='2010' style=''>2010</option><option value='2009' style=''>2009</option><option value='2008' style=''>2008</option><option value='2007' style=''>2007</option></select></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-text' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_976546488' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>2) Em que ano concluiu o ensino médio (2º grau)?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<input type='text' name='entry.976546488' value='' class='ss-q-short' id='entry_976546488' dir='auto' style='' />" & _
"</div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-select' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_1798122446' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>3) Este foi o seu primeiro curso superior?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<select name='entry.1798122446' id='entry_1798122446' style=''><option value='' style=''></option><option value='Sim' style=''>Sim</option><option value='Não' style=''>Não</option></select></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-text' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_375230756' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>3.1) Se não é o seu primeiro curso superior, quantos cursos já fez?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<input type='text' name='entry.375230756' value='' class='ss-q-short' id='entry_375230756' dir='auto' style='' />" & _
"</div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-text' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_694970404' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>3.2) Qual o último curso feito?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<input type='text' name='entry.694970404' value='' class='ss-q-short' id='entry_694970404' dir='auto' style='' />" & _
"</div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-grid' style='margin:12px 0;overflow-x:auto;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_1518199557' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>Ocupação" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<div>"

bloco2="<table border='0' cellpadding='5' cellspacing='0' style=''><thead><tr><td class='ss-gridnumbers ss-gridrow-leftlabel' style='text-align:left;border-bottom:1px solid #d3d8d3;min-width:100px;max-width:200px;padding-left:15px;'></td><td class='ss-gridnumbers' style='width: 33%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Sim</label></td><td class='ss-gridnumbers' style='width: 33%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Não</label></td></tr></thead><tbody><tr class='ss-gridrow ss-grid-row-odd' style='text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;background-color:#f2f2f2;'><td class='ss-gridrow ss-gridrow-leftlabel' style='text-align:left;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;min-width:100px;max-width:200px;padding-left:15px;'>4) Você está trabalhando atualmente?</td><td class='ss-gridrow' style='width: 33%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.249089749' value='Sim' id='group_249089749_1' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 33%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.249089749' value='Não' id='group_249089749_2' class='ss-q-radio' style='' /></div></td></tr><tr class='ss-gridrow ss-grid-row-even' style='text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;background-color:#fff;'><td class='ss-gridrow ss-gridrow-leftlabel' style='text-align:left;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;min-width:100px;max-width:200px;padding-left:15px;'>5) O seu trabalho é na mesma área de atuação do seu curso?</td><td class='ss-gridrow' style='width: 33%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.1065790069' value='Sim' id='group_1065790069_1' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 33%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.1065790069' value='Não' id='group_1065790069_2' class='ss-q-radio' style='' /></div></td></tr></tbody></table></div></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-paragraph-text' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_443472529' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>6) Por qual motivo escolheu fazer este curso?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<textarea name='entry.443472529' rows='8' cols='0' class='ss-q-long' id='entry_443472529' dir='auto' style='resize:vertical;width:70%;'></textarea>" & _
"</div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-select' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_1692168966' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>7) Se pudesse, gostaria de cursar mais 2 anos e obter um bacharelado?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'>Bacharelado é a graduação convencional com duração de 4 a 5 anos em média.</div>" & _
"<select name='entry.1692168966' id='entry_1692168966' style=''><option value='' style=''></option><option value='Sim' style=''>Sim</option><option value='Não' style=''>Não</option></select></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-paragraph-text' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_1384537123' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>7.1) Porque gostaria ou não de cursar mais dois anos e obter o bacharelado?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<textarea name='entry.1384537123' rows='8' cols='0' class='ss-q-long' id='entry_1384537123' dir='auto' style='resize:vertical;width:70%;'></textarea>" & _
"</div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-grid' style='margin:12px 0;overflow-x:auto;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_216120378' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>8) Qual a avaliação que você tem do curso, após a sua conclusão?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<div>"

bloco3="<table border='0' cellpadding='5' cellspacing='0' style=''><thead><tr><td class='ss-gridnumbers ss-gridrow-leftlabel' style='text-align:left;border-bottom:1px solid #d3d8d3;min-width:100px;max-width:200px;padding-left:15px;'></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Muito boa</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Boa</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Regular</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Ruim</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Muito ruim</label></td></tr></thead><tbody><tr class='ss-gridrow ss-grid-row-odd' style='text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;background-color:#f2f2f2;'><td class='ss-gridrow ss-gridrow-leftlabel' style='text-align:left;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;min-width:100px;max-width:200px;padding-left:15px;'></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.639646705' value='Muito boa' id='group_639646705_1' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.639646705' value='Boa' id='group_639646705_2' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.639646705' value='Regular' id='group_639646705_3' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.639646705' value='Ruim' id='group_639646705_4' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.639646705' value='Muito ruim' id='group_639646705_5' class='ss-q-radio' style='' /></div></td></tr></tbody></table></div></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-scale' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_2022993494' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>9) Em ordem de importância, sendo 1 o mais importante e 5 o menos, avalie os componentes do seu curso." & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>"
    
bloco4="<table border='0' cellpadding='5' cellspacing='0' id='entry_903595412' style=''><tbody><tr class='aria-todo' style=''><td class='ss-scalenumbers' style='text-align:center;'></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_903595412_1' style='display:block;padding:0.5em 0 .5em;'>1</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_903595412_2' style='display:block;padding:0.5em 0 .5em;'>2</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_903595412_3' style='display:block;padding:0.5em 0 .5em;'>3</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_903595412_4' style='display:block;padding:0.5em 0 .5em;'>4</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_903595412_5' style='display:block;padding:0.5em 0 .5em;'>5</label></td><td class='ss-scalenumbers' style='text-align:center;'></td></tr><tr><td class='ss-scalerow ss-leftlabel' style='text-align:right;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-left:0;'><div class='aria-todo' style=''>Conteúdo das disciplinas</div>" & _
"<div class='aria-only-help' style='font-size:0;left:-9999px;position:absolute;'>Selecione um valor no intervalo de 1,Conteúdo das disciplinas, a 5.</div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.903595412' value='1' id='group_903595412_1' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.903595412' value='2' id='group_903595412_2' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.903595412' value='3' id='group_903595412_3' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.903595412' value='4' id='group_903595412_4' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.903595412' value='5' id='group_903595412_5' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow ss-rightlabel aria-todo' style='text-align:left;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-right:0;'></td></tr></tbody></table></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-scale' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_1517035658' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>"

bloco5="<table border='0' cellpadding='5' cellspacing='0' id='entry_1689733854' style=''><tbody><tr class='aria-todo' style=''><td class='ss-scalenumbers' style='text-align:center;'></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1689733854_1' style='display:block;padding:0.5em 0 .5em;'>1</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1689733854_2' style='display:block;padding:0.5em 0 .5em;'>2</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1689733854_3' style='display:block;padding:0.5em 0 .5em;'>3</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1689733854_4' style='display:block;padding:0.5em 0 .5em;'>4</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1689733854_5' style='display:block;padding:0.5em 0 .5em;'>5</label></td><td class='ss-scalenumbers' style='text-align:center;'></td></tr><tr><td class='ss-scalerow ss-leftlabel' style='text-align:right;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-left:0;'><div class='aria-todo' style=''>Capacidade dos professores</div>" & _
"<div class='aria-only-help' style='font-size:0;left:-9999px;position:absolute;'>Selecione um valor no intervalo de 1,Capacidade dos professores, a 5.</div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1689733854' value='1' id='group_1689733854_1' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1689733854' value='2' id='group_1689733854_2' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1689733854' value='3' id='group_1689733854_3' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1689733854' value='4' id='group_1689733854_4' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1689733854' value='5' id='group_1689733854_5' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow ss-rightlabel aria-todo' style='text-align:left;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-right:0;'></td></tr></tbody></table></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-scale' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_1375525621' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>"

bloco6="<table border='0' cellpadding='5' cellspacing='0' id='entry_274305856' style=''><tbody><tr class='aria-todo' style=''><td class='ss-scalenumbers' style='text-align:center;'></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_274305856_1' style='display:block;padding:0.5em 0 .5em;'>1</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_274305856_2' style='display:block;padding:0.5em 0 .5em;'>2</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_274305856_3' style='display:block;padding:0.5em 0 .5em;'>3</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_274305856_4' style='display:block;padding:0.5em 0 .5em;'>4</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_274305856_5' style='display:block;padding:0.5em 0 .5em;'>5</label></td><td class='ss-scalenumbers' style='text-align:center;'></td></tr><tr><td class='ss-scalerow ss-leftlabel' style='text-align:right;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-left:0;'><div class='aria-todo' style=''>Coordenação do curso</div>" & _
"<div class='aria-only-help' style='font-size:0;left:-9999px;position:absolute;'>Selecione um valor no intervalo de 1,Coordenação do curso, a 5.</div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.274305856' value='1' id='group_274305856_1' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.274305856' value='2' id='group_274305856_2' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.274305856' value='3' id='group_274305856_3' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.274305856' value='4' id='group_274305856_4' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.274305856' value='5' id='group_274305856_5' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow ss-rightlabel aria-todo' style='text-align:left;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-right:0;'></td></tr></tbody></table></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-scale' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_893676728' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>"

bloco7="<table border='0' cellpadding='5' cellspacing='0' id='entry_1472994858' style=''><tbody><tr class='aria-todo' style=''><td class='ss-scalenumbers' style='text-align:center;'></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1472994858_1' style='display:block;padding:0.5em 0 .5em;'>1</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1472994858_2' style='display:block;padding:0.5em 0 .5em;'>2</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1472994858_3' style='display:block;padding:0.5em 0 .5em;'>3</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1472994858_4' style='display:block;padding:0.5em 0 .5em;'>4</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_1472994858_5' style='display:block;padding:0.5em 0 .5em;'>5</label></td><td class='ss-scalenumbers' style='text-align:center;'></td></tr><tr><td class='ss-scalerow ss-leftlabel' style='text-align:right;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-left:0;'><div class='aria-todo' style=''>Amadurecimento pessoal</div>" & _
"<div class='aria-only-help' style='font-size:0;left:-9999px;position:absolute;'>Selecione um valor no intervalo de 1,Amadurecimento pessoal, a 5.</div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1472994858' value='1' id='group_1472994858_1' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1472994858' value='2' id='group_1472994858_2' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1472994858' value='3' id='group_1472994858_3' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1472994858' value='4' id='group_1472994858_4' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.1472994858' value='5' id='group_1472994858_5' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow ss-rightlabel aria-todo' style='text-align:left;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-right:0;'></td></tr></tbody></table></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-scale' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_627081898' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>"

bloco8="<table border='0' cellpadding='5' cellspacing='0' id='entry_387422028' style=''><tbody><tr class='aria-todo' style=''><td class='ss-scalenumbers' style='text-align:center;'></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_387422028_1' style='display:block;padding:0.5em 0 .5em;'>1</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_387422028_2' style='display:block;padding:0.5em 0 .5em;'>2</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_387422028_3' style='display:block;padding:0.5em 0 .5em;'>3</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_387422028_4' style='display:block;padding:0.5em 0 .5em;'>4</label></td><td class='ss-scalenumbers' style='text-align:center;'><label class='ss-scalenumber' for='group_387422028_5' style='display:block;padding:0.5em 0 .5em;'>5</label></td><td class='ss-scalenumbers' style='text-align:center;'></td></tr><tr><td class='ss-scalerow ss-leftlabel' style='text-align:right;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-left:0;'><div class='aria-todo' style=''>Capacitação profissional</div>" & _
"<div class='aria-only-help' style='font-size:0;left:-9999px;position:absolute;'>Selecione um valor no intervalo de 1,Capacitação profissional, a 5.</div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.387422028' value='1' id='group_387422028_1' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.387422028' value='2' id='group_387422028_2' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.387422028' value='3' id='group_387422028_3' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.387422028' value='4' id='group_387422028_4' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow' style='text-align:center;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;'><div class='ss-scalerow-fieldcell' style=''><input type='radio' name='entry.387422028' value='5' id='group_387422028_5' class='ss-q-radio' style='' /></div></td><td class='ss-scalerow ss-rightlabel aria-todo' style='text-align:left;color:#666;border:1px solid #d3d8d3;border-left:0;border-right:0;padding:.5em .25em;padding-right:0;'></td></tr></tbody></table></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-grid' style='margin:12px 0;overflow-x:auto;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_2035971750' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>10) Em termos de aplicabilidade, o que o curso trouxe de capacitação para a sua carreira?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<div>"

bloco9="<table border='0' cellpadding='5' cellspacing='0' style=''><thead><tr><td class='ss-gridnumbers ss-gridrow-leftlabel' style='text-align:left;border-bottom:1px solid #d3d8d3;min-width:100px;max-width:200px;padding-left:15px;'></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Muito boa</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Boa</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Regular</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Ruim</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Muito ruim</label></td></tr></thead><tbody><tr class='ss-gridrow ss-grid-row-odd' style='text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;background-color:#f2f2f2;'><td class='ss-gridrow ss-gridrow-leftlabel' style='text-align:left;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;min-width:100px;max-width:200px;padding-left:15px;'>Capacitação</td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.902671815' value='Muito boa' id='group_902671815_1' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.902671815' value='Boa' id='group_902671815_2' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.902671815' value='Regular' id='group_902671815_3' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.902671815' value='Ruim' id='group_902671815_4' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.902671815' value='Muito ruim' id='group_902671815_5' class='ss-q-radio' style='' /></div></td></tr></tbody></table></div></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-grid' style='margin:12px 0;overflow-x:auto;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_1177390451' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>11) Qual a imagem que você tem da instituição de ensino que escolheu?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<div>"

bloco10="<table border='0' cellpadding='5' cellspacing='0' style=''><thead><tr><td class='ss-gridnumbers ss-gridrow-leftlabel' style='text-align:left;border-bottom:1px solid #d3d8d3;min-width:100px;max-width:200px;padding-left:15px;'></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Muito boa</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Boa</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Regular</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Ruim</label></td><td class='ss-gridnumbers' style='width: 16%;text-align:center;border-bottom:1px solid #d3d8d3;'><label class='ss-gridnumber' style='display:block;padding:0.5em 0 .5em;'>Muito ruim</label></td></tr></thead><tbody><tr class='ss-gridrow ss-grid-row-odd' style='text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;background-color:#f2f2f2;'><td class='ss-gridrow ss-gridrow-leftlabel' style='text-align:left;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;min-width:100px;max-width:200px;padding-left:15px;'>Imagem da instituição</td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.575466795' value='Muito boa' id='group_575466795_1' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.575466795' value='Boa' id='group_575466795_2' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.575466795' value='Regular' id='group_575466795_3' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.575466795' value='Ruim' id='group_575466795_4' class='ss-q-radio' style='' /></div></td><td class='ss-gridrow' style='width: 16%;text-align:center;color:#666;border-bottom:1px solid #d3d8d3;padding:.5em .25em;'><div class='ss-grid-button-wrapper' style=''><input type='radio' name='entry.575466795' value='Muito ruim' id='group_575466795_5' class='ss-q-radio' style='' /></div></td></tr></tbody></table></div></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-radio' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_1812478504' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>12) O que fez após concluir seu curso?" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>"

bloco11="<ul class='ss-choices' style='list-style:none;margin:.5em 0 0 0;padding:0;'><li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.650485752' value='Cursou uma pós-graduação' id='group_650485752_1' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>Cursou uma pós-graduação</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.650485752' value='Procurou outra oportunidade de trabalho menor' id='group_650485752_2' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>Procurou outra oportunidade de trabalho menor</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.650485752' value='Mudou a área de atuação' id='group_650485752_3' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>Mudou a área de atuação</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.650485752' value='Parou de estudar' id='group_650485752_4' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>Parou de estudar</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.650485752' value='Fez outro curso' id='group_650485752_5' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>Fez outro curso</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.650485752' value='Empreendeu um negócio próprio' id='group_650485752_6' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>Empreendeu um negócio próprio</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.650485752' value='Estudou uma língua estrangeira' id='group_650485752_7' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>Estudou uma língua estrangeira</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.650485752' value='__other_option__' id='group_650485752_8' class='ss-q-radio ss-q-other-toggle' style='' /></span>" & _
"Outro:</label>" & _
"<span class='ss-q-other-container goog-inline-block' style='position:relative;display:inline-block;'><input type='text' name='entry.650485752.other_option_response' value='' class='ss-q-other' id='entry_650485752_other_option_response' dir='auto' style='' /></span></li></ul></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-radio' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_2082555450' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>13) Faixa etária" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" 

bloco12="<ul class='ss-choices' style='list-style:none;margin:.5em 0 0 0;padding:0;'><li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.764162949' value='Até 20 anos' id='group_764162949_1' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>Até 20 anos</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.764162949' value='De 20 a 25 anos' id='group_764162949_2' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>De 20 a 25 anos</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.764162949' value='De 26 a 30 anos' id='group_764162949_3' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>De 26 a 30 anos</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.764162949' value='De 31 a 35 anos' id='group_764162949_4' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>De 31 a 35 anos</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.764162949' value='De 36 a 40 anos' id='group_764162949_5' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>De 36 a 40 anos</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.764162949' value='De 41 a 45 anos' id='group_764162949_6' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>De 41 a 45 anos</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.764162949' value='De 46 a 50 anos' id='group_764162949_7' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>De 46 a 50 anos</span>" & _
"</label></li> <li class='ss-choice-item' style='margin:0;line-height:1.3em;padding-bottom:.5em;'><label><span class='ss-choice-item-control goog-inline-block' style='position:relative;display:inline-block;'><input type='radio' name='entry.764162949' value='Acima de 50 anos' id='group_764162949_8' class='ss-q-radio' style='' /></span>" & _
"<span class='ss-choice-label' style=''>Acima de 50 anos</span>" & _
"</label></li></ul></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-select' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_991768653' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>14) Sexo" & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<select name='entry.991768653' id='entry_991768653' style=''><option value='' style=''></option><option value='Feminino' style=''>Feminino</option><option value='Masculino' style=''>Masculino</option></select></div></div></div> <div class='ss-form-question errorbox-good' style=''>" & _
"<div dir='ltr' class='ss-item  ss-text' style='margin:12px 0;'><div class='ss-form-entry' style='max-width:100%;'><label class='ss-q-item-label' for='entry_629366060' style=''></label><div class='ss-q-title' style='display:block;font-weight:bold;'>Caso deseje receber os resultados desta pesquisa, informe o seu email." & _
"</div>" & _
"<div class='ss-q-help ss-secondary-text' dir='ltr' style='display:block;margin:.1em 0 .25em 0;color:#666;'></div>" & _
"<input type='text' name='entry.629366060' value='' class='ss-q-short' id='entry_629366060' dir='auto' style='' />" & _
"</div></div></div>" & _
"<input type='hidden' name='draftResponse' value='[]' style='' />" & _
"<input type='hidden' name='pageHistory' value='0' style='' />"

bloco13="<div class='ss-item ss-navigate' style='margin:12px 0;clear:both;'><div class='ss-form-entry' style='max-width:100%;'>" & _
"<input type='submit' name='submit' value='Enviar' id='ss-submit' style='' />" & _
"<div class='ss-secondary-text' style='color:#666;'>Nunca envie senhas em formulários do Google.</div></div></div></form></div>" & _
"<div class='ss-footer' style=''><div class='ss-attribution' style=''></div>" & _
"<div class='ss-legal' style=''><div class='disclaimer-separator' style=''></div>" & _
"<div class='disclaimer' style=''><div class='powered-by-logo' style='margin-top:2em;'><span class='powered-by-text' style=''>Powered by</span>" & _
"<a class='ss-logo-link' href='http://docs.google.com' style='display:inline-block;text-decoration:none;' target='_blank'><img class='ss-logo' src='https://ssl.gstatic.com/docs/forms/drive_logo_small.png' alt='Google Drive' style='border:none;height:23px;width:105px;' /></a></div>" & _
"<div class='ss-terms' style='color:#777;font-size:11px;margin-top:1.5em;'><span class='disclaimer-msg' style=''>Este conteúdo não foi criado nem aprovado pelo Google.</span>" & _
"<br />" & _
"<a href='https://docs.google.com/forms/d/1pALiC_5H732umOqohpIewFQAv_FcThQrjzU8jS-Elo0/reportabuse?source=https://docs.google.com/forms/d/1pALiC_5H732umOqohpIewFQAv_FcThQrjzU8jS-Elo0/viewform?sid%3D5dd59ee26950313a%26token%3Dskk48j4BAAA.lrPcJY3AjM5JPdcHRbfUgw.Pb08MVnleDQqQH4GbEJWUg' style='' target='_blank'>Denunciar abuso</a>" & _
"-" & _
"<a href='http://www.google.com/accounts/TOS' style='' target='_blank'>Termos de Serviço</a>" & _
"-" & _
"<a href='http://www.google.com/google-d-s/terms.html' style='' target='_blank'>Termos Adicionais</a></div></div></div></div>" & _
"</div></div>" 

	texto2="</td></tr></table>"
	'response.write texto2
	
	geral=cabecalho & intro & texto & bloco1 & bloco2 & bloco3 & bloco4 & bloco5 & bloco6 & bloco7 & bloco8 & bloco9 & bloco10 & bloco11 & bloco12 & bloco13 & texto2
	'response.write geral

	Set Mailer = CreateObject("CDO.Message") 
	Mailer.From = "Pesquisa Unifieo <pesquisa.unifieo@gmail.com>" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = email ' e-mail de quem vai receber a mensagem 
	'Mailer.CC = "02379@unifieo.br" ' Com Cópia 
	'Mailer.BCC = "00259@unifieo.br" ' Com Cópia 
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "Pesquisa sobre conclusão de curso tecnológico - Unifieo" 
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=geral
	
	'"<br><br><p style='margin-bottom:0;margin-top:15'><b>Recursos Humanos</b><br>3651-9905"
'==This section provides the configuration information for the remote SMTP server.
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2   '2 
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "pesquisa.unifieo@gmail.com"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "eb541627"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
Mailer.Configuration.Fields.Update
'==End remote SMTP server configuration section==
	response.write "<br>" & rs("mataluno") & " - Enviando email para: " & email
	teste=1
	if teste=1 then Mailer.Send 
	Set Mailer = Nothing 

rs.movenext
loop
response.write "</table>"

end if
rs.close

teste2=1
if teste2=1 then
%>
<%
end if

%>
	
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>