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
<title>Formul�rio para Transfer�ncia de Funcion�rio</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr>
	<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
	<td><b><input type="text" value="AUTORIZA��O PARA TRANSFER�NCIA DE FUNCION�RIO" size=50 class=form_input10 style="font-weight:bold;"></b><td>
</tr>
</table>
<br><br>
<table cellpadding="5" cellspacing="0" width="650" style="border:1px solid #000000">
    <tr><td class="campop">Entrevistado por: <input type="text" value="" size=50 class=form_input10></td></tr>
	<tr><td class="campop">Processo Seletivo: <input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Nome do Funcion�rio</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Departamento atual: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Departamento para transfer�ncia: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000" width=300>
	<i>Local de trabalho atual: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Local de trabalho para transfer�ncia: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Cargo atual: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Cargo na transfer�ncia: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Hor�rio atual: </i></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Hor�rio na transfer�ncia: </i></td>
	<td class=campo style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000" width=250>
	<i>Sal�rio atual: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Sal�rio na transfer�ncia: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Data de Transfer�ncia</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" name="admissao" class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Entrevistado(a) pela Chefia:</i>&nbsp;<input type="text" value="" size=64 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Coment�rios:</i>&nbsp;<input type="text" value="" size=76 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size=88 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>


	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="radio" name="aprov1" value="A"> Aprovado &nbsp;
 	<input type="radio" name="aprov1" value="N"> N�o aprovado &nbsp; &nbsp;
	<input type="text" value="" size=35 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Recursos Humanos:</td>
<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Pr�-Reitoria Administrativa:</td></tr>

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