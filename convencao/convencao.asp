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
<title>Pontos importantes da Conven��o Coletiva 2005</title>
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
<tr><td valign=top class=titulo colspan=3>Pontos importantes das Conven��es Coletivas 2005</td></tr>
<tr><td valign=top class=grupo colspan=3>Auxiliares</td></tr>
<tr><td valign=top class=titulo>Cl�usula</td>
	<td valign=top class=titulo>Teor atual</td>
	<td valign=top class=titulo>Teor anterior</td></tr>

<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Dura��o</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Mar�o/2005 a Fev/2007</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Mar�o/2004 a Fev/2005</td></tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Reajuste</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000"
	>Negociado em 7,66%.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000"
	>Negociado em 7,48%</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Para 2006, pr�-definido 
	pela m�dia aritm�tica do INPC, IPC e ICVMar�o/2005 a Fev/2007.<br>
	Se o c�lculo ultrapassar 9,99%, vai ser dado 9,99% e o que passar vai ser negociado.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">-</td></tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Abono de Faltas por Casamento ou Luto</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">�9 dias corridos por 
	casamento ou luto de pai, m�e, filho(a), c�mjuge, companheiro(a) ou dependente juridicamente reconhecido.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">�9 dias corridos por 
	casamento ou luto de pai, m�e, filho(a), c�mjuge, companheiro(a) ou dependente juridicamente reconhecido.</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000"><font color=blue>�3 dias no 
	caso de falecimento de irm�o(a), sogro(a) e neto(a).</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">-</td></tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Bolsas de Estudo</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	�Tem direito <font color=blue>ap�s a experi�ncia.</font></td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	�Tem direito a qualquer tempo.</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">�Tem direito al�m do 
	Auxiliar, esposo(a) e companheiro(a), <font color=blue>filhos e dependentes sob guarda</font> at� 25 anos.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">�Tem direito al�m do Auxiliar, 
	filhos e esposas at� 24 anos.</td></tr>

<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Prazo para Homologa��o</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Deve homologar at� o 
	20� dia ap�s o pagamento das verbas rescis�rias.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Deve homologar em at� 20 
	dias ap�s o t�rmino do aviso pr�vio trabalhado, ou<br>em at� 30 dias ap�s o desligamento quando o aviso � indenizado.</td></tr>

<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Assist�ncia M�dica</td>
	<td valign=top class="campot" colspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000;border-right: 1px solid #000000">
	Mediante o pagamento de 10% da mensalidade
	da assist�ncia (Cl�ssico I=R$ 6,31) o Auxiliar pode continuar como benefici�rio ap�s sua sa�da da empresa de acordo com a 
	Lei 9656/98.<br><br>
	<i>Art.30 - Ao consumidor que contribuir para plano ou seguro privado coletivo de assist�ncia � sa�de, decorrente de v�nculo empregat�cio,
	no caso de rescis�o ou exonera��o do contrato de trabalho sem justa causa, � assegurado o direito de manter sua condi��o de
	benefici�rio, nas mesmas condi��es de que gozava quando da vig�ncia do contrato de trabalho, desde que assuma tamb�m o pagamento
	da parcela anteriormente de responsabilidade patronal.<br>
	� 1� <font color=blue>O per�odo de manuten��o</font> da condi��o de benefici�rio a que se refere o caput ser� 
	de <font color=blue>um ter�o (1/3) do tempo de perman�ncia</font> no plano ou seguro, ou sucessor, com um <font color=blue>m�nimo</font> 
	de <font color=blue>seis (6) meses</font> e um <font color=blue>m�ximo</font> de <font color=blue>vinte e quatro (24) meses</font>.<br>
	� 2� A manuten��o de que trata este artigo <font color=blue>� extensiva, obrigatoriamente, a todo o grupo familiar inscrito</font> quando da vig�ncia do
	contrato de trabalho.<br>
	� 3� Em caso de morte do titular, o direito de perman�ncia � assegurado aos dependentes cobertos pelo plano ou seguro privado 
	coletivo de assist�ncia � sa�de, nos termos do disposto neste artigo.<br>
	� 4� O direito assegurado neste artigo n�o exclui vantagens obtidas pelos empregados decorrentes de negocia��es coletivas de trabalho.</i>
	<br>
	<br>O artigo 31 da mesma lei fala sobre o aposentado, que tenha contribuido por pelo menos 10 anos. Ao se afastar, pode manter o 
	benef�cio, a raz�o de 1 ano para cada ano de contribui��o.
	</td>
	</tr>
	
<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000"><b>Piso Salarial</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">R$ 490,92 por jornada de 44 hs/semanais</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">R$ 456,00 por jornada de 44 hs/semanais</td></tr>
	
<tr><td valign=top class="campol" style="border: 1px solid #000000;border-bottom: 2 solid #000000"><b>Cesta B�sica</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Tem direito quem recebe at� 5 sal�rios m�nimos por m�s, 
	<font color=blue>em jornada integral de 44 horas semanais</font>.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">Tem direito quem recebe at� 5 sal�rios m�nimos por m�s.</td></tr>
</table>

<br>

<!-- professores -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td valign=top class=grupo colspan=3>Professores</td></tr>
<tr><td valign=top class=titulo>Cl�usula</td>
	<td valign=top class=titulo>Teor atual</td>
	<td valign=top class=titulo>Teor anterior</td></tr>

<tr><td valign=top class="campol" colspan=3 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Cl�usulas como <b>Dura��o</b>, <b>Reajuste</b>, <b>Assist�ncia M�dica</b> s�o id�nticas �s do Auxiliar</td>
	</tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Janela</td>
	<td valign=top class="campot" colspan=2 style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	O pagamento da Janela � obrigat�rio, ressalvada a condi��o quando o Professor aceita atrav�s de acordo formalizado antes do in�cio
	das aulas, n�o ser pago pelas janelas.</td></tr>
	<tr><td valign=top class="campot" colspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">Ocorrendo este acordo, caso o 
	Professor venha a ministrar aulas (substitui��o, etc) ou exercer outra atividade (est�gio, TCC, etc), no hor�rio da janela n�o
	paga, estas atividades ser�o remuneradas como aulas extras, com adicional de 100%.</td></tr>

<tr><td valign=top class="campol" rowspan=1 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Dura��o da Hora-Aula</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">A dura��o da hora-aula poder�
	ser de no m�ximo 50 minutos, <font color=blue>salvo nos Curso Tecnol�gicos que tenham sido autorizados ou reconhecidos com dura��o da hora-aula
	de 60 minutos, e que os professores destes cursos tenham sido contratados nesta condi��o.</font></td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">A dura��o da hora-aula poder�
	ser de no m�ximo 50 minutos.</td></tr>

<tr><td valign=top class="campol" rowspan=1 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Redu��o de Carga Hor�ria por Supress�o ou Extin��o de Disciplina</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">O professor dever� ser comunicado da
	redu��o da carga hor�ria com <font color=blue>anteced�ncia de 30 dias do in�cio do per�odo letivo.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">N�o existia a cl�usula.</td></tr>
	
<tr><td valign=top class="campol" rowspan=1 style="border: 1px solid #000000;border-bottom:2 solid #000000">
	<b>Redu��o de Carga Hor�ria por Supress�o de Curso, Turma ou Disciplina</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">O professor dever� ser comunicado da
	redu��o da carga hor�ria at� <font color=blue>o final da 2� semana do per�odo de aulas..</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">mesma descri��o</td></tr>
	
<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Abono de Faltas por Casamento ou Luto</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">�9 dias corridos por 
	casamento ou luto de pai, m�e, filho(a), c�mjuge, companheiro(a) ou dependente juridicamente reconhecido.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">�9 dias corridos por 
	casamento ou luto de pai, m�e, filho(a), c�mjuge, companheiro(a) ou dependente juridicamente reconhecido.</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000"><font color=blue>�Aten��o!! N�o existe
	outros abonos para o SINPRO, somente para a Federa��o. Irm�o(a) continua valendo a CLT, 2 dias.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">-</td></tr>

<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Garantia Semestral de Sal�rios</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	�Tem direito quem est� h� pelo menos 18 meses na institui��o.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">�id�ntico</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	�Nas demiss�es a partir de 16/outubro, independente do tempo de servi�o, haver� indeniza��o at� 18/janeiro,</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">�id�ntico</td></tr>

	
<tr><td valign=top class="campol" rowspan=2 style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Bolsas de Estudo</td>
	<td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">
	�Tem direito a qualquer tempo.</font></td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 dotted #000000">�id�ntico.</td></tr>
	<tr><td valign=top class="campot" style="border: 1px solid #000000;border-bottom: 2 solid #000000">�Tem direito al�m do 
	Professor, filhos e dependentes legais. Filhos at� 25 anos.</td>
	<td valign=top class="campoa" style="border: 1px solid #000000;border-bottom: 2 solid #000000">�id�ntico.</td></tr>

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