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
<title>Pontos importantes da Conven��o Coletiva 2007</title>
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
<tr><td valign=top class=titulo colspan=3>Pontos importantes das Conven��es Coletivas 2008/9</td></tr>
<tr><td valign=top class=grupo colspan=3>Auxiliares</td></tr>
<tr><td valign=top class=titulo width=10%>Cl�usula</td>
	<td valign=top class=titulo width=45%>Teor atual</td>
	<td valign=top class=titulo width=45%>Teor anterior</td></tr>

<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	29. Indeniza��o por Dispensa Imotivada
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Par�grafo terceiro</b> � O pagamento das verbas indenizat�rias previstas nesta cl�usula n�o ser� cumulativo, cabendo ao AUXILIAR, no desligamento, o maior valor monet�rio entre os previstos nas al�neas �a� e �b� do caput.
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	N�o existia.<font style="text-decoration:line-through;font-weight:bold"></font>
	<blockquote style="margin-top:0;margin-bottom:0">a) 03 (tr�s) dias para cada ano trabalhado na MANTENEDORA;
	<br>b) aviso pr�vio adicional de quinze dias, caso o AUXILIAR tenha, no m�nimo, cinq�enta anos de idade e que, � data do desligamento, conte com pelo menos um ano de servi�o na MANTENEDORA.
	</blockquote>
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	41. Assist�ncia M�dico-hospitalar
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	5. Pagamento � A assist�ncia m�dico-hospitalar ser� garantida nos termos desta Conven��o, cabendo ao AUXILIAR, para usufruir dos benef�cios da Lei n� 9656/98, o pagamento de 10% das mensalidades da referida assist�ncia, respeitado o estabelecido no par�grafo 1� (primeiro) desta cl�usula.
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	5. Pagamento � A assist�ncia m�dico-hospitalar ser� garantida nos termos desta Conven��o, cabendo ao AUXILIAR, para usufruir dos benef�cios da Lei n� 9656/98, o pagamento de 10% das mensalidades da referida assist�ncia, <font style="text-decoration:line-through;font-weight:bold">com teto limite de R$ 8,00 (oito reais) por m�s,</font> respeitado o estabelecido no par�grafo 1� (primeiro) desta cl�usula.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	43. Menor sal�rio da categoria
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	R$ 561,63 (at� Fev/2009)<br>
	R$ 595,32 (Mai/2009)<br>
	R$ 603,19 (a partir de Jun/2009)
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	R$ 529,80
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	49. Cesta B�sica
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada aos AUXILIARES que percebam, <b>at� 4 (quatro) vezes o piso salarial da categoria</b>, em jornada integral de 44 (quarenta e quatro) horas semanais, ou percebam, em jornada inferior, remunera��o proporcionalmente igual ou inferior ao limite fixado nesta cl�usula, a concess�o de uma cesta b�sica mensal de 26 kg
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Fica assegurada aos AUXILIARES que percebam, <b>at� 5 (cinco) sal�rios m�nimos por m�s</b>, em jornada de 36 (trinta e seis) horas semanais, ou percebam, em jornada inferior, remunera��o proporcionalmente igual ou inferior ao limite fixado nesta cl�usula, a concess�o de uma cesta b�sica mensal de 26 kg
	</td></tr>




</table>

<DIV style="page-break-after:always"></DIV>

<!-- professores -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td valign=top class=titulo colspan=3>Pontos importantes das Conven��es Coletivas 2008/9</td></tr>
<tr><td valign=top class=grupo colspan=3>Professores</td></tr>
<tr><td valign=top class=titulo width=10%>Cl�usula</td>
	<td valign=top class=titulo width=45%>Teor atual</td>
	<td valign=top class=titulo width=45%>Teor anterior</td></tr>

<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	1. Abrang�ncia
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo segundo - exclu�do 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	Par�grafo segundo � Quando o PROFESSOR for contratado em um munic�pio para exercer a sua atividade em outro, prevalecer� o cumprimento da Conven��o Coletiva do munic�pio onde o servi�o � prestado.
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	24. Irredutibilidade Salarial
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Par�grafo terceiro</b> � A MANTENEDORA n�o poder� reduzir o valor da hora-aula dos contratos de trabalho vigentes, ainda que venha a instituir ou modificar plano de carreira. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	30. Pedido de demiss�o em final de ano letivo
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	O PROFESSOR que, no final do ano letivo, comunicar sua demiss�o at� o dia que antecede o in�cio do recesso escolar, ser� dispensado do cumprimento do aviso pr�vio e ter� direito a receber, como indeniza��o, a remunera��o at� o dia 18 de janeiro do ano subseq�ente, independentemente do tempo de servi�o na MANTENEDORA. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	39. F�rias
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Par�grafo terceiro</b> � O pagamento das verbas indenizat�rias previstas nesta cl�usula n�o ser� cumulativo, cabendo ao PROFESSOR, no desligamento, o maior valor monet�rio entre os previstos nas al�neas a) e b) do caput. 
	<br><b>Par�grafo quarto</b> � Essas indeniza��es n�o contar�o, para nenhum efeito, como tempo de servi�o. 	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	&nbsp;
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	50. Assist�ncia M�dica
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>5. Pagamento </b>
	<br>Caber� ao PROFESSOR o pagamento de 10% (dez por cento) do valor da Assist�ncia M�dica, respeitado o disposto nos par�grafos 1�, 2� e 3�. 
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>5. Pagamento </b>
	<br>Caber� ao PROFESSOR o pagamento de 10% (dez por cento) do valor da Assist�ncia M�dica, <font style="text-decoration:line-through;font-weight:bold">limitado tal pagamento a R$ 8,00</font>, respeitado o disposto nos par�grafos 1� e 2�. 
	</td></tr>
<tr><td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	50. Assist�ncia M�dica
	</td>
	<td valign=top class=campo style="border: 1px solid #000000;border-bottom: 2 solid #000000">
	<b>Par�grafo primeiro</b> � A MANTENEDORA dever� enviar ao SINPRO c�pia do contrato formalizado com a empresa de assist�ncia m�dico�hospitalar ou de seguro sa�de ou de medicina de grupo que comprove o valor pago. 
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