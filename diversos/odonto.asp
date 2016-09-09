<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Pesquisa de interesse para adesão de Plano Odontológico</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<div align="right">
<table border="0" style="border-collapse: collapse" cellpadding="5" width="690" height="990" cellspacing="0">
<tr><td width="100%" valign="top" height="20"><hr style="border:dotted 1 #000000"></td></tr>
<%for b=1 to 2%>
<tr><td width="100%" valign="top" height="450"> <!-- celula para definir tamanho-->
<!-- inicio formulario -->
<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td height=30 align="left" class="campop" style="border-bottom:solid 2 #000000">
	<b>Pesquisa de interesse para adesão de Plano Odontológico</td>
	<td align="right" valign=middle style="border-bottom:solid 2 #000000">
	<img src="../images/logo_centro_universitario_unifieo_big.jpg" width="150" border="0" alt="">
	</td>
</tr>
<tr><td height=15></td></tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:2;margin-bottom:2;line-height:15px;text-align:justify"><b>
	Preencha o formulário abaixo e devolva ao departamento de Recursos Humanos.<br>
	Contamos com sua colaboração.
	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	1 - Alguma vez você recebeu orientação sobre a importância da odontologia para a sua saúde geral e sobre
	Prevenção em Odontologia?
	</td>
</tr>
<tr>
	<td class="campop">
	<input type="radio" name="perg1" value="Sim"> Sim &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="radio" name="perg1" value="Não"> Não
	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	2 - Você já tem plano odontológico?
	</td>
</tr>
<tr>
	<td class="campop">
	<input type="radio" name="perg2" value="Sim"> Sim &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="radio" name="perg2" value="Não"> Não
	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	3 - Caso o UNIFIEO oferecesse um Plano Coletivo Odontológico para você e seus dependentes, com ampla cobertura para
	restaurações, tratamentos de canal, tratamentos gengivais, emergência 24 horas, odontopediatria, prevenção, entre
	outros em consultórios e clínicas particulares, você teria interesse em participar?
	</td>
</tr>
<tr>
	<td class="campop">
	<input type="radio" name="perg3" value="Sim"> Sim &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="radio" name="perg3" value="Não"> Não
	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	4 - Pelo serviço oferecido você aceitaria pagar entre R$ 10,00 e R$ 12,00 mensais por pessoa a ser descontado
	no seu holerite?
	</td>
</tr>
<tr>
	<td class="campop">
	<input type="radio" name="perg3" value="Sim"> Sim &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="radio" name="perg3" value="Não"> Não
	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop" colspan=3><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	5 - O que motivaria a sua participação em um plano odontológico?
	</td>
</tr>
<tr>
	<td class="campop">
	<input type="checkbox" name="perg5" value="Preço"> Preço<br>
	<input type="checkbox" name="perg5" value="Rede Credenciada Ampla"> Rede Credenciada Ampla<br>
	<input type="checkbox" name="perg5" value="Cobertura Ampla"> Cobertura Ampla<br>
	<input type="checkbox" name="perg5" value="Outros"> Outros<br>
	</td>
	<td class="campop" valign="middle" align="center"> Quais?</td>
	<td class="campop" valign="top" align="right">
		<input type="text" name="perg5" size="75" style="height:80px">
	</td>
</tr>
</table>

<!-- celula fim para definir tamanho -->	
</td></tr>
<tr><td width="100%" valign="top" height="20"><hr style="border:dotted 1 #000000"></td></tr>
<%next%>
	
</table>
</div>
</body>
</html>