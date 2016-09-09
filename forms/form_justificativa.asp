<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Justificativa para Aus�ncia de Marca��o de Ponto</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<div align="right">
<table border="0" style="border-collapse: collapse" cellpadding="5" width="650" height="1000" cellspacing="0">
<tr><td width="100%" valign="top" height="0"></td></tr>
<%for b=1 to 2%>
<tr><td width="100%" valign="top" >
	<table style="border-collapse: collapse" border="1" bordercolor="#CCCCCC" cellpadding="2" width="640" cellspacing="0">
	<tr>
		<td><img border="0" src="../images/logo_centro_universitario_unifieo_big.jpg" width="225" height="50"></td>
		<td align="center"><b><font size="2">Justificativa para Aus�ncia de Marca��o de Ponto</font></b></td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="640" cellspacing="0">
	<tr>
		<td valign="top"><font size="1">Departamento:</font><br>&nbsp;</td>
		<td width="150" valign="top"><font size="1">M�s:</font><br>&nbsp;</td>
		<td width="100" valign="top"><font size="1">Ano:</font><br><b><%=year(now())%></b></td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="640" cellspacing="0">
	<tr>
		<td width="80" valign="top"><font size="1">Chapa:</font><br>&nbsp;</td>
		<td valign="top"><font size="1">Nome do Funcion�rio:</font><br>&nbsp;</td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="640" cellspacing="0">
	<tr>
		<td width="100%" valign="top" colspan="12"><font size="1">Destina-se o presente controle a registrar informa��es do Empregado,
		relativas aos dias e hor�rio de trabalho face a justificativa assinalada. Fica ciente o empregado, e autoriza, que as informa��es ser�o
		<u>inclu�das</u> manualmente nas suas marca��es de ponto e conferidas com outros controles eletr�nicos dispon�veis, como catraca eletr�nica, controle de estacionamento, entre outros.</font></td></tr>
	<tr>
		<td class=fundor colspan=5 align="center" style="border:1px solid #000000"><i><b>Informe apenas as marca��es a serem inclu�das</td>
		<td class=fundor colspan=7 align="center" style="border:1px solid #000000"><i><b>Assinale o motivo</td>
	</tr>
	<tr>
		<td width="30" valign="middle" rowspan="2" align="center"><font size="1">DIA</font></td>
		<td width="60" valign="middle" rowspan="2" align="center"><font size="1">Hor�rio de Entrada</font></td>
		<td            valign="top"    colspan="2" align="center"><font size="1">Intervalo para refei��o</font></td>
		<td width="60" valign="middle" rowspan="2" align="center"><font size="1">Hor�rio de Sa�da</font></td>
		<td            valign="middle" colspan="7" align="center"><font size="1">Justificativa p/ Aus�ncia</font></td>
	</tr>
	<tr>
		<td width="60" valign="top" align="center"><font size="1">Sa�da</font></td>
		<td width="60" valign="top" align="center"><font size="1">Retorno</font></td>
		<td width="20" valign="top" align="center"><font size="1">EM</font></td>
		<td width="20" valign="top" align="center"><font size="1">EC</font></td>
		<td width="20" valign="top" align="center"><font size="1">AM</font></td>
		<td width="20" valign="top" align="center"><font size="1">TE</font></td>
		<td width="20" valign="top" align="center"><font size="1">SP</font></td>
		<td width="20" valign="top" align="center"><font size="1">RD</font></td>
		<td width="210"valign="top" align="center"><font size="1">Outros</font></td>
	</tr>
<%for a=1 to 6%>
	<tr>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
	</tr>
<%next%>
	<tr>
		<td valign="top" colspan="12"><font size="1">C�d. Justificativas: &nbsp;
		<b>EM</b> - Esquecimento de marca��o | 
		<b>EC</b> - Esquecimento do crach� |
		<b>AM</b> - Apagar marca��es em excesso/duplicidade<br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<b>TE</b> - Trabalho Externo <b><i>(anexar relat�rio identificando)</i></b> |
		<b>SP</b> - Sem Papel |
		<b>RD</b> - Rel�gio desligado 
		</font></td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="640" cellspacing="0">
	<tr>
		<td width="100%">
		<table style="border-collapse: collapse"  border="0" cellpadding="0" width="100%">
		<tr>
			<td width="25%" class="campor" valign="bottom">&nbsp;<br>_____________________<br>Data</td>
			<td width="25%" class="campor" valign="bottom">&nbsp;<br>__________________________<br>Assinatura do Funcion�rio</td>
			<td width="50%" class="campor" valign="bottom" style="border-left:1px solid #000000"><b>Confirmo a legitimidade das informa��es desta justificativa:</b><br><br>&nbsp;&nbsp;&nbsp;___________________________________<br>&nbsp;&nbsp;&nbsp;Assinatura da Chefia</td>
		</tr></table>
		</td>
	</tr></table>

	<table border="0" cellpadding="2" width="640" cellspacing="0">
	<tr><td width="100%" align="right" class="campor">Form.RH 02/2016</td>
	</tr></table>

	</td>
</tr>
<tr><td width="100%" valign="top" height="0"></td></tr>
<%next%>
	
</table>
</div>
</body>
</html>