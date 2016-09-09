<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a51")="N" or session("a51")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Formulário para Transferência de Funcionário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr>
	<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
	<td><b><input type="text" value="AUTORIZAÇÃO PARA TRANSFERÊNCIA DE FUNCIONÁRIO" size=50 class=form_input10 style="font-weight:bold;"></b><td>
</tr>
</table>
<br><br>
<table cellpadding="5" cellspacing="0" width="650" style="border:1px solid #000000">
    <tr><td class="campop">Entrevistado por: <input type="text" value="" size=50 class=form_input10></td></tr>
	<tr><td class="campop">Processo Seletivo: <input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Nome do Funcionário</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Departamento atual: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Departamento para transferência: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000" width=300>
	<i>Local de trabalho atual: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Local de trabalho para transferência: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Cargo atual: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Cargo na transferência: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Horário atual: </i></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Horário na transferência: </i></td>
	<td class=campo style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000" width=250>
	<i>Salário atual: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Salário na transferência: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Data de Transferência</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" name="admissao" class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
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