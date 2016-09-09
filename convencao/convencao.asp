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
<title>Pontos importantes da Convenção Coletiva 2005</title>
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
<tr><td valign=top class=titulo colspan=3>Pontos importantes das Convenções Coletivas 2005</td></tr>
<tr><td valign=top class=grupo colspan=3>Auxiliares</td></tr>
<tr><td valign=top class=titulo>Cláusula</td>
	<td valign=top class=titulo>Teor atual</td>
	<td valign=top class=titulo>Teor anterior</td></tr>

<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Duração</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Março/2005 a Fev/2007</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Março/2004 a Fev/2005</td></tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Reajuste</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000"
	>Negociado em 7,66%.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000"
	>Negociado em 7,48%</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Para 2006, pré-definido 
	pela média aritmética do INPC, IPC e ICVMarço/2005 a Fev/2007.<br>
	Se o cálculo ultrapassar 9,99%, vai ser dado 9,99% e o que passar vai ser negociado.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">-</td></tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Abono de Faltas por Casamento ou Luto</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">•9 dias corridos por 
	casamento ou luto de pai, mãe, filho(a), cômjuge, companheiro(a) ou dependente juridicamente reconhecido.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">•9 dias corridos por 
	casamento ou luto de pai, mãe, filho(a), cômjuge, companheiro(a) ou dependente juridicamente reconhecido.</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000"><font color=blue>•3 dias no 
	caso de falecimento de irmão(a), sogro(a) e neto(a).</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">-</td></tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Bolsas de Estudo</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	•Tem direito <font color=blue>após a experiência.</font></td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	•Tem direito a qualquer tempo.</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">•Tem direito além do 
	Auxiliar, esposo(a) e companheiro(a), <font color=blue>filhos e dependentes sob guarda</font> até 25 anos.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">•Tem direito além do Auxiliar, 
	filhos e esposas até 24 anos.</td></tr>

<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Prazo para Homologação</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Deve homologar até o 
	20º dia após o pagamento das verbas rescisórias.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Deve homologar em até 20 
	dias após o término do aviso prévio trabalhado, ou<br>em até 30 dias após o desligamento quando o aviso é indenizado.</td></tr>

<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Assistência Médica</td>
	<td valign=top class="campot" colspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000;border-right: 1px solid #000000">
	Mediante o pagamento de 10% da mensalidade
	da assistência (Clássico I=R$ 6,31) o Auxiliar pode continuar como beneficiário após sua saída da empresa de acordo com a 
	Lei 9656/98.<br><br>
	<i>Art.30 - Ao consumidor que contribuir para plano ou seguro privado coletivo de assistência à saúde, decorrente de vínculo empregatício,
	no caso de rescisão ou exoneração do contrato de trabalho sem justa causa, é assegurado o direito de manter sua condição de
	beneficiário, nas mesmas condições de que gozava quando da vigência do contrato de trabalho, desde que assuma também o pagamento
	da parcela anteriormente de responsabilidade patronal.<br>
	§ 1º <font color=blue>O período de manutenção</font> da condição de beneficiário a que se refere o caput será 
	de <font color=blue>um terço (1/3) do tempo de permanência</font> no plano ou seguro, ou sucessor, com um <font color=blue>mínimo</font> 
	de <font color=blue>seis (6) meses</font> e um <font color=blue>máximo</font> de <font color=blue>vinte e quatro (24) meses</font>.<br>
	§ 2º A manutenção de que trata este artigo <font color=blue>é extensiva, obrigatoriamente, a todo o grupo familiar inscrito</font> quando da vigência do
	contrato de trabalho.<br>
	§ 3º Em caso de morte do titular, o direito de permanência é assegurado aos dependentes cobertos pelo plano ou seguro privado 
	coletivo de assistência à saúde, nos termos do disposto neste artigo.<br>
	§ 4º O direito assegurado neste artigo não exclui vantagens obtidas pelos empregados decorrentes de negociações coletivas de trabalho.</i>
	<br>
	<br>O artigo 31 da mesma lei fala sobre o aposentado, que tenha contribuido por pelo menos 10 anos. Ao se afastar, pode manter o 
	benefício, a razão de 1 ano para cada ano de contribuição.
	</td>
	</tr>
	
<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000"><b>Piso Salarial</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">R$ 490,92 por jornada de 44 hs/semanais</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">R$ 456,00 por jornada de 44 hs/semanais</td></tr>
	
<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000"><b>Cesta Básica</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Tem direito quem recebe até 5 salários mínimos por mês, 
	<font color=blue>em jornada integral de 44 horas semanais</font>.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Tem direito quem recebe até 5 salários mínimos por mês.</td></tr>
</table>

<br>

<!-- professores -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td valign=top class=grupo colspan=3>Professores</td></tr>
<tr><td valign=top class=titulo>Cláusula</td>
	<td valign=top class=titulo>Teor atual</td>
	<td valign=top class=titulo>Teor anterior</td></tr>

<tr><td valign=top class="campol" colspan=3 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Cláusulas como <b>Duração</b>, <b>Reajuste</b>, <b>Assistência Médica</b> são idênticas às do Auxiliar</td>
	</tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Janela</td>
	<td valign=top class="campot" colspan=2 style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	O pagamento da Janela é obrigatório, ressalvada a condição quando o Professor aceita através de acordo formalizado antes do início
	das aulas, não ser pago pelas janelas.</td></tr>
	<tr><td valign=top class="campot" colspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">Ocorrendo este acordo, caso o 
	Professor venha a ministrar aulas (substituição, etc) ou exercer outra atividade (estágio, TCC, etc), no horário da janela não
	paga, estas atividades serão remuneradas como aulas extras, com adicional de 100%.</td></tr>

<tr><td valign=top class="campol" rowspan=1 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Duração da Hora-Aula</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">A duração da hora-aula poderá
	ser de no máximo 50 minutos, <font color=blue>salvo nos Curso Tecnológicos que tenham sido autorizados ou reconhecidos com duração da hora-aula
	de 60 minutos, e que os professores destes cursos tenham sido contratados nesta condição.</font></td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">A duração da hora-aula poderá
	ser de no máximo 50 minutos.</td></tr>

<tr><td valign=top class="campol" rowspan=1 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Redução de Carga Horária por Supressão ou Extinção de Disciplina</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">O professor deverá ser comunicado da
	redução da carga horária com <font color=blue>antecedência de 30 dias do início do período letivo.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Não existia a cláusula.</td></tr>
	
<tr><td valign=top class="campol" rowspan=1 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Redução de Carga Horária por Supressão de Curso, Turma ou Disciplina</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">O professor deverá ser comunicado da
	redução da carga horária até <font color=blue>o final da 2ª semana do período de aulas..</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">mesma descrição</td></tr>
	
<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Abono de Faltas por Casamento ou Luto</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">•9 dias corridos por 
	casamento ou luto de pai, mãe, filho(a), cômjuge, companheiro(a) ou dependente juridicamente reconhecido.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">•9 dias corridos por 
	casamento ou luto de pai, mãe, filho(a), cômjuge, companheiro(a) ou dependente juridicamente reconhecido.</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000"><font color=blue>•Atenção!! Não existe
	outros abonos para o SINPRO, somente para a Federação. Irmão(a) continua valendo a CLT, 2 dias.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">-</td></tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Garantia Semestral de Salários</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	•Tem direito quem está há pelo menos 18 meses na instituição.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">•idêntico</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	•Nas demissões a partir de 16/outubro, independente do tempo de serviço, haverá indenização até 18/janeiro,</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">•idêntico</td></tr>

	
<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Bolsas de Estudo</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	•Tem direito a qualquer tempo.</font></td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">•idêntico.</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">•Tem direito além do 
	Professor, filhos e dependentes legais. Filhos até 25 anos.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">•idêntico.</td></tr>

</table>

<%
'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>
<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>