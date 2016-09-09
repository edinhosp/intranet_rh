<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Comunicação de Falta, Atraso ou Saída Antecipada</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<div align="right">
<table border="0" style="border-collapse: collapse" cellpadding="5" width="650" height="990" cellspacing="0">
<tr><td width="100%" valign="top" height="30"><hr></td></tr>
<%for b=1 to 2%>
<tr><td width="100%" valign="top" height="450"> <!-- celula para definir tamanho-->
<!-- inicio formulario -->
<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="600">
<tr>
	<td height=38 align="center" class="campop" style="border:double 5px #000000">
	<b>COMUNICAÇÃO DE FALTA, ATRASO OU SAÍDA ANTECIPADA</td>
</tr>
<tr><td height=15></td></tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="2" cellspacing="0" width="600">
<tr><td colspan=3 align="left" class="campop" style="border-bottom:solid 1px #000000">
	<b>Nome:</td></tr>
<tr><td height=5></td></tr>
<tr><td height=38 class="campop" align="left" style="border:solid 1px #000000">
	&nbsp;Faltou no dia _____/_____ ou no período de _____/_____ a _____/_____</td></tr>
<tr><td height=5></td></tr>
</table>
	
<table style="border-collapse: collapse" border="0" cellpadding="5" cellspacing="0" width="600">
<tr>
	<td height=70 style="border:ridge 1px #CCCCCC">
		<table style="border-collapse: collapse" border="0" cellpadding="2" cellspacing="0">
		<tr><td height=30 valign=middle colspan=2>Atrasou no dia _____/_____/_____</td></tr>
		<tr>
		<td width=80 style="border:solid 1px #000000">HORA DE<br>ENTRADA</td>
		<td width=90 style="border:solid 1px #000000">&nbsp;</td>
		</tr>
		</table>
	</td>
	<td width=1></td>
	<td height=70 style="border:ridge 1 #CCCCCC">
		<table width=344 style="border-collapse: collapse" border="0" cellpadding="2" cellspacing="0">
		<tr><td height=30 valign=middle colspan=5>Retirou-se no dia _____/_____/_____</td></tr>
		<tr>
		<td width=80 style="border:solid 1px #000000">HORA DE<br>SAÍDA</td>
		<td width=90 style="border:solid 1px #000000">&nbsp;</td>
		<td width=1></td>
		<td width=80 style="border:solid 1px #000000">HORA DE<br>RETORNO</td>
		<td width=90 style="border:solid 1px #000000">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>		
</table>

<table style="border-collapse: collapse" border="0" cellpadding="2" cellspacing="0" width="600">
<tr><td colspan=13 height=30 align="left" class="campop" style="border-bottom:solid 1px #000000">
	<b>Motivo:</td></tr>
<tr><td height=5 colspan=13></td></tr>
<tr><td>Autorizado</td>
	<td></td>
	<td>SIM</td>
	<td></td>
	<td width=20 style="border:solid 1px #000000">&nbsp;</td>	
	<td></td>
	<td>NÃO</td>
	<td></td>
	<td width=20 style="border:solid 1px #000000">&nbsp;</td>	
	<td width=5></td>
	<td width=90></td>
	<td width=5></td>
	<td width=200></td>
</tr>
<tr><td height=5 colspan=13></td></tr>
<tr><td>Anexa Comprovante</td>
	<td></td>
	<td>SIM</td>
	<td></td>
	<td style="border:solid 1px #000000">&nbsp;</td>	
	<td></td>
	<td>NÃO</td>
	<td></td>
	<td style="border:solid 1px #000000">&nbsp;</td>	
	<td></td>
	<td style="border-bottom:solid 1px #000000"></td>
	<td></td>
	<td style="border-bottom:solid 1px #000000"></td>
</tr>
<tr><td colspan=9></td>
	<td></td>
	<td width=90>Data</td>
	<td></td>
	<td width=200>Empregado</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="2" cellspacing="0" width="600">
<tr>
	<td class="campor" rowspan=3 valign=top align="left" style="border:solid 1px #000000">INSTRUÇÕES AO RECURSOS HUMANOS</td>
	<td></td>
	<td class="campor" style="border:solid 1px #000000">ABONAR</td>
	<td width=20 style="border:solid 1 #000000">&nbsp;</td>
	<td></td>
	<td class="campor" style="border:solid 1px #000000">DESCONTAR<br>DIA/HORA</td>
	<td width=20 style="border:solid 1 #000000">&nbsp;</td>
	<td></td>
	<td class="campor" style="border:solid 1px #000000">DESCONTAR<br>DIA/HORA + DSR</td>
	<td width=20 style="border:solid 1 #000000">&nbsp;</td>
</tr>
<tr><td colspan=9 height=5></td></tr>
<tr>
	<td></td>
	<td height=35 colspan=8 valign=top class=campo style="border:solid 1px #000000">Obs.:</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="2" cellspacing="0" width="600">
<tr><td colspan=3 height=5></td></tr>
<tr>
	<td height=40 valign=top align="left" style="border:solid 1px #000000">Superior Imediato</td>
	<td></td>
	<td valign=top align="left" style="border:solid 1px #000000">Data</td>
</tr>

	
</table>	
<!-- celula fim para definir tamanho -->	
</td></tr>
<tr><td width="100%" valign="top" height="30"><hr></td></tr>
<%next%>
	
</table>
</div>
</body>
</html>