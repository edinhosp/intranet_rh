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
		<td class="campop" align="center"><b>Relatório de Horas Prestadas</b></td>
	</tr></table>

	<table bordercolor=#000000 style="border-collapse: collapse"  border="1" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td class="campor" valign=top>Local da Prestação:<br>&nbsp;</td>
		<td class="campor" width=150 valign=top>Mês:<br>&nbsp;</td>
		<td class="campor" width=100 valign=top>Ano:<br><b><%=year(now())%></b></td>
	</tr></table>

	<table bordercolor=#000000 style="border-collapse: collapse"  border="1" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td class="campor" width=80 valign=top>Código:<br>&nbsp;</td>
		<td class="campor" valign=top>Nome do Prestador:<br>&nbsp;</td>
	</tr></table>
		  
	<table bordercolor=#000000 style="border-collapse: collapse"  border="1" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td class="campor" valign=top colspan=7>Destina-se o presente controle a registrar informações do Prestador, relativas aos dias e horário da prestação de serviços.</td></tr>
	<tr>
		<td class="campor" width=30 valign=middle rowspan=2 align="center">DIA                    </td>
		<td class="campor" width=60 valign=middle rowspan=2 align="center">Horário de Entrada     </td>
		<td class="campor"            valign=top    colspan=2 align="center">Intervalo para refeição</td>
		<td class="campor" width=60 valign=middle rowspan=2 align="center">Horário de Saída      </td>
		<td class="campor" width=50 valign="center" rowspan=2 align="center">Total do dia          </td>
		<td class="campor" width=280 valign=middle rowspan=2 align="center">Descrição dos serviços</td>
	</tr>
	<tr>
		<td class="campor" width=60 valign=top align="center">Saída  </td>
		<td class="campor" width=60 valign=top align="center">Retorno</td>
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
		<td class="campor" valign=top height=22>&nbsp;</td>
	</tr>
<%
next
%>
	<tr>
		<td class="campor" valign=top height=22 colspan=5>&nbsp;Total de Horas do Período</td>
		<td class="campor" valign=top height=22>&nbsp;</td>
		<td class="campor" valign=top height=22>&nbsp;</td>
	</tr>
	</table>

	<table bordercolor=#000000 style="border-collapse: collapse"  border="1" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="100%">
		<table style="border-collapse: collapse"  border="0" cellpadding="0" width="100%"><tr>
			<td class="campor" width="40%" valign=top>Assinatura do Prestador<br>&nbsp;<br>&nbsp;</td>
			<td class="campor" width="40%" valign=top>Assinatura do Coordenador<br>&nbsp;<br>&nbsp;</td>
			<td class="campor" width="20%" valign=top>Assinatura do Supervisor<br>&nbsp;</td>
			</tr></table>
		</td>
	</tr></table>

	<table border="0" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="100%" align="right">
			<p style="margin-top: 0; margin-bottom: 0"><font size="1">Form.RH 12/2005</font></td>
	</tr>
	</table>

	</td>
	</tr>
	</table>
</div>
</body>
</html>