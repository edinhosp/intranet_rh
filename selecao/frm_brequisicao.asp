<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a49")="N" or session("a49")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Formul�rio para Abertura de Vaga</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr>
	<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
	<td><b><input type="text" value="FORMUL�RIO PARA ABERTURA DE VAGA" size=50 class=form_input10 style="font-weight:bold;"></b><td>
</tr>
</table>
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>FUN��O/CARGO</i></td></tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="80" class=form_input10>&nbsp;</b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>SE��O/DEPT�</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>CAMPUS</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="40" class=form_input10>&nbsp;</b></td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="40" class=form_input10>&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>REQUISITANTE</i></td></tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="80" class=form_input10>&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;" valign="top">
	<i>Motivo da<br>Contrata��o</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000;">
	<input type="radio" name="o1" value="02"> Substitui��o<br>
	<input type="radio" name="o1" value="03"> Nova Vaga<br>
	<input type="radio" name="o1" value="04"> Aumento de quadro
	</td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000" valign="top">
	<i>se substitui��o, informar nome do Substitu�do</i><br>
	<input type="text" value="" size="60" class=form_input10>&nbsp;
	</td>
<tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000" valign="top">
	<i>Tipo do<br>Contratado</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="radio" name="o2" value="1"> Normal<br>
	<input type="radio" name="o2" value="2"> Estagi�rio
	</td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000" valign="top">
	<i>Faixa Salarial</i><br>
	<input type="text" value="" size="20" class=form_input10>&nbsp;
	</td>

	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000" valign="top">
	<i>Sal�rio Contrata��o / Efetiva��o</i><br>
	<input type="text" value="" size="20" class=form_input10>&nbsp;
	</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Hor�rio de Trabalho</i></td>
	<td class="campop" width=20% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Jornada Mensal</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="70" class=form_input10>&nbsp;</td>
	<td class="campop" style="border-right:1px solid #000000" valign=top>
	<input type="text" value="" size="5" class=form_input10> horas</td>
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Benef�cios oferecidos</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="90" class=form_input10>&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<i>Escolaridade, Forma��o ou Curso T�cnico exigido</i><br>
	<input type="text" value="" size="60" class=form_input10>
	</td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<input type="radio" name="o3" value="1"> prefer�vel<br>
	<input type="radio" name="o3" value="2"> exigida
	</td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Experi�ncia M�nima</i><br>
	<input type="text" value="" size="5" class=form_input10> anos
	</td>
</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;" valign="top">
	<i>Sexo</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000" valign="top">
	<input type="radio" name="os1" value="I"> Indiferente<br>
	<input type="radio" name="os1" value="F"> Feminino<br>
	<input type="radio" name="os1" value="M"> Masculino
	</td>	
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000" valign="top">
	<i>Idade</i><br>
	m�nima:<input type="text" value="" size="5" class=form_input10>&nbsp;&nbsp;m�xima:<input type="text" value="" size="5" class=form_input10>
	</td>
	<td class="campop" style="border-top:1px solid #000000;" valign="top">
	<i>Deficiente<br>Tipo def.</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000" valign="top">
	<input type="radio" name="od1" value="0"> Indiferente<br>
	<input type="radio" name="od1" value="1"> N�o deficiente<br>
	<input type="radio" name="od1" value="2"> Deficiente <input type="text" value="" size="15" class=form_input10>
	</td>
<tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Conhecimentos espec�ficos</i><br>
	<input type="text" value="" size="80" class=form_input10>&nbsp;
	</td>
</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" height=85>
<tr><td height=15 class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Descri��es das atividades</i><br>
	><input type="text" value="" size="45" class=form_input10>&nbsp;<br>
	><input type="text" value="" size="45" class=form_input10>&nbsp;<br>
	><input type="text" value="" size="45" class=form_input10>&nbsp;
	</td>
	<td height=15 class="campop" style="border-top:1px solid #000000;;border-right:1px solid #000000">
	<i>Responsabilidades</i><br>
	><input type="text" value="" size="45" class=form_input10>&nbsp;<br>
	><input type="text" value="" size="45" class=form_input10>&nbsp;<br>
	><input type="text" value="" size="45" class=form_input10>&nbsp;
	</td>	
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Data de Abertura</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Data de Encerramento</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Quant. Vagas dispon�veis</i></td></tr>

<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size="20" class=form_input10>&nbsp;&nbsp;</td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="20" class=form_input10>&nbsp;&nbsp;</td>
	<td class="campop" style="border-right:1px solid #000000">
	<input type="text" value="" size="20" class=form_input10>&nbsp;&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>&nbsp;</i></td></tr>
	<tr><td class="campop" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td></tr>
</table>
<%for a=1 to 4%>
<br>
<%next%>
<p align="center">Recursos Humanos</p>
</body>
</html>