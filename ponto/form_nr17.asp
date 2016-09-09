<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Controle de Presença e Horas de Trabalho</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<div align="right">
<table style="border-collapse: collapse"  border="0" cellpadding="5" width="650" height="990" cellspacing="0">
<tr>
	<td valign="top" height="450">
	<table style="border-collapse: collapse" border="0" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td><img border="0" src="../images/logo_centro_universitario_unifieo_big.gif" width="225" height="50"></td>
		<td class="campop" align="center"><b>Lançamento de Pausas - Portaria 09/2007-MTE</b></td>
	</tr></table>

	<table bordercolor=#000000 style="border-collapse: collapse"  border="1" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td class="campor" valign=top>Local de Trabalho:<br>&nbsp;</td>
		<td class="campor" width=150 valign=top>Mês:<br>&nbsp;</td>
		<td class="campor" width=100 valign=top>Ano:<br><b><%=year(now())%></b></td>
	</tr></table>

	<table bordercolor=#000000 style="border-collapse: collapse"  border="1" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td class="campor" width=80 valign=top>Chapa:<br>&nbsp;</td>
		<td class="campor" valign=top>Nome do Funcionário:<br>&nbsp;</td>
	</tr></table>
		  
	<table bordercolor=#000000 style="border-collapse: collapse"  border="1" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td class="campor" valign=top colspan=6>Destina-se o presente controle a registrar informações das 2 (duas) pausas diárias de
		10 minutos instituídas pela Portaria 09/2007-MTE, Anexo II da NR17, item 5.4.</td></tr>
	<tr>
		<td class=fundor width=30 valign=middle align="center"><b>DIA</td>
		<td class=fundor width=100 valign=middle align="center"><b>Horário da 1ª Pausa</td>
		<td class=fundor width=80 valign="center" align="center"><b>Término</td>
		<td class=fundor width=100 valign=middle align="center"><b>Horário da 2ª Pausa</td>
		<td class=fundor width=80 valign=middle align="center"><b>Término</td>
		<td class=fundor width=210 valign=middle align="center"></td>
	</tr>
<%
for a=1 to 31
%>
	<tr>
		<td class="campor" valign=top height=22 align="center">&nbsp;<%=a%></td>
		<td class="campor" valign=top height=22>&nbsp;</td>
		<td class="campor" valign=top height=22>&nbsp;</td>
		<td class="campor" valign=top height=22>&nbsp;</td>
		<td class="campor" valign=top height=22>&nbsp;</td>
		<td class="campor" valign=top height=22>&nbsp;</td>
	</tr>
<%
next
%>
	</table>

	<table bordercolor=#000000 style="border-collapse: collapse"  border="1" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="100%">
		<table style="border-collapse: collapse"  border="0" cellpadding="0" width="100%"><tr>
			<td class="campor" width="40%" valign=top>Assinatura do Estagiário<br>&nbsp;<br>&nbsp;</td>
			<td class="campor" width="40%" valign=top>Assinatura do Coordenador<br>&nbsp;<br>&nbsp;</td>
			<td class="campor" width="20%" valign=top>Assinatura do Supervisor<br>&nbsp;</td>
			</tr></table>
		</td>
	</tr></table>

	<table border="0" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="100%" align="right">
			<p style="margin-top: 0; margin-bottom: 0"><font size="1">Form.RH 12/2013</font></td>
	</tr>
	</table>

	</td>
	</tr>
	</table>
</div>
</body>
</html>