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
<title>Pontos importantes da Convenção Coletiva 2007</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<!-- -->
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
<tr><td valign=top class=titulo colspan=3>Pontos importantes das Convenções Coletivas 2008/9</td></tr>
<tr><td valign=top class=grupo colspan=3>Auxiliares</td></tr>
<tr><td valign=top class=titulo width=10%>Cláusula</td>
	<td valign=top class=titulo width=45%>Teor atual</td>
	<td valign=top class=titulo width=45%>Teor anterior</td></tr>

<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	29. Indenização por Dispensa Imotivada
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Parágrafo terceiro</b> – O pagamento das verbas indenizatórias previstas nesta cláusula não será cumulativo, cabendo ao AUXILIAR, no desligamento, o maior valor monetário entre os previstos nas alíneas “a” e “b” do caput.
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Não existia.<font style="text-decoration:line-through;font-weight:bold"></font>
	<blockquote style="margin-top:0;margin-bottom:0">a) 03 (três) dias para cada ano trabalhado na MANTENEDORA;
	<br>b) aviso prévio adicional de quinze dias, caso o AUXILIAR tenha, no mínimo, cinqüenta anos de idade e que, à data do desligamento, conte com pelo menos um ano de serviço na MANTENEDORA.
	</blockquote>
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	41. Assistência Médico-hospitalar
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	5. Pagamento – A assistência médico-hospitalar será garantida nos termos desta Convenção, cabendo ao AUXILIAR, para usufruir dos benefícios da Lei nº 9656/98, o pagamento de 10% das mensalidades da referida assistência, respeitado o estabelecido no parágrafo 1º (primeiro) desta cláusula.
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	5. Pagamento – A assistência médico-hospitalar será garantida nos termos desta Convenção, cabendo ao AUXILIAR, para usufruir dos benefícios da Lei nº 9656/98, o pagamento de 10% das mensalidades da referida assistência, <font style="text-decoration:line-through;font-weight:bold">com teto limite de R$ 8,00 (oito reais) por mês,</font> respeitado o estabelecido no parágrafo 1º (primeiro) desta cláusula.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	43. Menor salário da categoria
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	R$ 561,63 (até Fev/2009)<br>
	R$ 595,32 (Mai/2009)<br>
	R$ 603,19 (a partir de Jun/2009)
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	R$ 529,80
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	49. Cesta Básica
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada aos AUXILIARES que percebam, <b>até 4 (quatro) vezes o piso salarial da categoria</b>, em jornada integral de 44 (quarenta e quatro) horas semanais, ou percebam, em jornada inferior, remuneração proporcionalmente igual ou inferior ao limite fixado nesta cláusula, a concessão de uma cesta básica mensal de 26 kg
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada aos AUXILIARES que percebam, <b>até 5 (cinco) salários mínimos por mês</b>, em jornada de 36 (trinta e seis) horas semanais, ou percebam, em jornada inferior, remuneração proporcionalmente igual ou inferior ao limite fixado nesta cláusula, a concessão de uma cesta básica mensal de 26 kg
	</td></tr>




</table>

<DIV style="page-break-after:always"></DIV>

<!-- professores -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td valign=top class=titulo colspan=3>Pontos importantes das Convenções Coletivas 2008/9</td></tr>
<tr><td valign=top class=grupo colspan=3>Professores</td></tr>
<tr><td valign=top class=titulo width=10%>Cláusula</td>
	<td valign=top class=titulo width=45%>Teor atual</td>
	<td valign=top class=titulo width=45%>Teor anterior</td></tr>

<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	1. Abrangência
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo segundo - excluído 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Parágrafo segundo – Quando o PROFESSOR for contratado em um município para exercer a sua atividade em outro, prevalecerá o cumprimento da Convenção Coletiva do município onde o serviço é prestado.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	24. Irredutibilidade Salarial
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Parágrafo terceiro</b> – A MANTENEDORA não poderá reduzir o valor da hora-aula dos contratos de trabalho vigentes, ainda que venha a instituir ou modificar plano de carreira. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	30. Pedido de demissão em final de ano letivo
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	O PROFESSOR que, no final do ano letivo, comunicar sua demissão até o dia que antecede o início do recesso escolar, será dispensado do cumprimento do aviso prévio e terá direito a receber, como indenização, a remuneração até o dia 18 de janeiro do ano subseqüente, independentemente do tempo de serviço na MANTENEDORA. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	39. Férias
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Parágrafo terceiro</b> – O pagamento das verbas indenizatórias previstas nesta cláusula não será cumulativo, cabendo ao PROFESSOR, no desligamento, o maior valor monetário entre os previstos nas alíneas a) e b) do caput. 
	<br><b>Parágrafo quarto</b> – Essas indenizações não contarão, para nenhum efeito, como tempo de serviço. 	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	50. Assistência Médica
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>5. Pagamento </b>
	<br>Caberá ao PROFESSOR o pagamento de 10% (dez por cento) do valor da Assistência Médica, respeitado o disposto nos parágrafos 1º, 2º e 3º. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>5. Pagamento </b>
	<br>Caberá ao PROFESSOR o pagamento de 10% (dez por cento) do valor da Assistência Médica, <font style="text-decoration:line-through;font-weight:bold">limitado tal pagamento a R$ 8,00</font>, respeitado o disposto nos parágrafos 1º e 2º. 
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	50. Assistência Médica
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Parágrafo primeiro</b> – A MANTENEDORA deverá enviar ao SINPRO cópia do contrato formalizado com a empresa de assistência médico–hospitalar ou de seguro saúde ou de medicina de grupo que comprove o valor pago. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
</table>

<%
'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>
<!-- -->
<!-- -->
</body>
</html>