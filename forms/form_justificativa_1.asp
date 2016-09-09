<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Justificativa para Ausência de Marcação de Ponto</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<div align="right">
<table border="0" style="border-collapse: collapse" cellpadding="5" width="650" height="990" cellspacing="0">
<tr><td width="100%" valign="top" height="30"><hr></td></tr>
<%for b=1 to 2%>
<tr><td width="100%" valign="top" height="450">
	<table style="border-collapse: collapse" border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td><img border="0" src="../images/logo_centro_universitario_unifieo_big.jpg" width="225" height="50"></td>
		<td align="center"><b><font size="2">Justificativa para Ausência de Marcação de Ponto</font></b></td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td valign="top"><font size="1">Departamento:</font><br>&nbsp;</td>
		<td width="150" valign="top"><font size="1">Mês:</font><br>&nbsp;</td>
		<td width="100" valign="top"><font size="1">Ano:</font><br><b><%=year(now())%></b></td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="80" valign="top"><font size="1">Chapa:</font><br>&nbsp;</td>
		<td valign="top"><font size="1">Nome do Funcionário:</font><br>&nbsp;</td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="100%" valign="top" colspan="8"><font size="1">Destina-se o presente controle a registrar informações do Empregado,
		relativas aos dias e horário de trabalho face a justificativa assinamada. Fica ciente o empregado, e autoriza, que as informações serão
		incluídas manualmente nas suas marcações de ponto e conferidas com outros controles eletrônicos disponíveis, como catraca eletrônica, controle de estacionamento, entre outros.</font></td></tr>
	<tr>
		<td width="30" valign="middle" rowspan="2" align="center"><font size="1">DIA</font></td>
		<td width="60" valign="middle" rowspan="2" align="center"><font size="1">Horário de Entrada</font></td>
		<td width="80" valign="middle" rowspan="2" align="center"><font size="1">Assinatura do funcionário</font></td>
		<td            valign="top"    colspan="2" align="center"><font size="1">Intervalo para refeição</font></td>
		<td width="60" valign="middle" rowspan="2" align="center"><font size="1">Horário de Saída</font></td>
		<td width="80" valign="middle" rowspan="2" align="center"><font size="1">Assinatura do funcionário</font></td>
		<td            valign="middle" rowspan="2" align="center"><font size="1">Justificativa p/ Ausência</font></td>
	</tr>
	<tr>
		<td width="60" valign="top" align="center"><font size="1">Saída</font></td>
		<td width="60" valign="top" align="center"><font size="1">Retorno</font></td>
	</tr>
<%for a=1 to 6%>
	<tr>
		<td width="30" valign="top" height="25">&nbsp;</td>
		<td width="60" valign="top" height="25">&nbsp;</td>
		<td width="80" valign="top" height="25">&nbsp;</td>
		<td width="60" valign="top" height="25">&nbsp;</td>
		<td width="60" valign="top" height="25">&nbsp;</td>
		<td width="60" valign="top" height="25">&nbsp;</td>
		<td width="80" valign="top" height="25">&nbsp;</td>
		<td            valign="top" height="25">&nbsp;</td>
	</tr>
<%next%>
	<tr>
		<td valign="top" colspan="8"><font size="1">Cód. Justificativas:<br>
		1019 - Esquecimento de marcação&nbsp; 1020 - Esquecimento do
		cartão 1027 - Serviço externo&nbsp; 1028 - Prob. Téc. Equipamento</font></td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="100%">
		<table style="border-collapse: collapse"  border="0" cellpadding="0" width="100%">
		<tr>
			<td width="33%"></td>
			<td width="33%">&nbsp;<br>_____________________<br><font size="1">Data</font></td>
			<td width="34%">&nbsp;<br>__________________________<br><font size="1">Assinatura da Chefia</font></td>
		</tr></table>
		</td>
	</tr></table>

	<table border="0" cellpadding="2" width="600" cellspacing="0">
	<tr><td width="100%" align="right"><p style="margin-top:0;margin-bottom:0"><font size="1">Form.RH 09/2003</font></td>
	</tr></table>

	</td>
</tr>
<tr><td width="100%" valign="top" height="30"><hr></td></tr>
<%next%>
	
</table>
</div>
</body>
</html>