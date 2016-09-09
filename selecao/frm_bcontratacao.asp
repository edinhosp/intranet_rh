<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a50")="N" or session("a50")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Autorização para Contratação de Funcionário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr>
	<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
	<td><b><input type="text" value="AUTORIZAÇÃO PARA CONTRATAÇÃO DE FUNCIONÁRIO" size=50 class=form_input10 style="font-weight:bold;">
	</b><td>
			 
</tr>
</table>
<br><br>

<table cellpadding="5" cellspacing="0" width="650" style="border:1px solid #000000">
    <tr><td class="campop">Entrevistado por: <input type="text" value="" size=50 class=form_input10></td></tr>
	<tr><td class="campop">Processo Seletivo: <input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Departamento Requisitante</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10>&nbsp;&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Nome do Candidato</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10>&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Cargo</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10>&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Local de Trabalho</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	Campus</b><input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Horário de Trabalho</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Jornada Mensal</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="70" class=form_input10>&nbsp;</td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="3" class=form_input10>&nbsp;horas</td>	
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td colspan=2 class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Salário</i></td></tr>
	<tr>
		<td class="campop" style="border-left:1px solid #000000" width=50%>
	inicial: <input type="text" value="" size="15" class=form_input10></td>
		<td class="campop" style="border-right:1px solid #000000" width=50%>
	após experiência: <input type="text" value="" size="15" class=form_input10></td>
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Data de Admissão</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" name="admissao" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-left:1px solid #000000">
	<input type="radio" name="motivo" value="02"> Substituição<br>
 	<input type="radio" name="motivo" value="03"> Vaga Nova<br>
 	<input type="radio" name="motivo" value="04"> Aumento de Quadro
	</td>
	<td class="campop" style="border-right:1px solid #000000" valign=top>
	<input type="text" value="" size="50" class=form_input10>&nbsp;</td>
	</tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000">
	<i>Motivo da Admissão</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-left: 1px solid;border-right:1px solid #000000" colspan=2>
	<i>Nome do Substituído</i></td>
	</tr>

	<tr><td class="campop" style="border-left:1px solid #000000" rowspan=4>
	<input type="radio" name="motivo" value="02"> Substituição<br>
 	<input type="radio" name="motivo" value="03"> Vaga Nova<br>
 	<input type="radio" name="motivo" value="04"> Aumento de Quadro
	</td>
	<td class="campop" style="border-right:1px solid #000000;border-left: 1px solid" valign=top colspan=2>
	<input type="text" value="" size="50" class=form_input10></td>
	</tr>
	<tr>
		<td class="campop" style="border-left: 1px solid;border-right:1px solid #000000;border-top: 1px solid" valign=top><i>Salário do substituido</td>
		<td class="campop" style="border-right: 1px solid;border-top: 1px solid" valign=top><i>Jornada do substituido</td>
	</tr>
	<tr>
		<td class="campop" style="border-lefT: 1px solid;border-right:1px solid #000000" valign=top><input type="text" value="" size="10" class=form_input10></td>
		<td class="campop" style="border-right:1px solid #000000" valign=top><input type="text" value="" size="3" class=form_input10> horas</td>
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Entrevistado(a) pela Chefia:</i>&nbsp;<input type="text" value="" size=64 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Comentários:</i>&nbsp;<input type="text" value="" size=76 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size=88 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>


	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="radio" name="aprov1" value="A"> Aprovado &nbsp;
 	<input type="radio" name="aprov1" value="N"> Não aprovado &nbsp; &nbsp;
	<input type="text" value="" size=35 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Recursos Humanos:</td>
<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Pró-Reitoria Administrativa:</td></tr>

	<tr><td class="campop" height=50 style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td>
<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td></tr>

</table>

<%for a=1 to 4%>
<br>
<%next%>

</body>
</html>