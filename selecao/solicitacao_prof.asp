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
<title>Solicita��o de Emprego - Professor</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<%
espacamento=5
if request("tipo")="C" then
	titulo1="Professor Convidado"
	titulo2="Cadastro"
else
	titulo1="Solicita��o de Emprego"
	titulo2="Cadastro de Docente"
end if
%>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop" width=240 rowspan=3 valign=top align="left"><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td>
	<td class="campop" width=410 height=71 align="center"><p style="font-family:'Century Gothic';font-size:18pt;margin-top:0;margin-bottom:0">
	<b><%=titulo1%></b><br><%=titulo2%></td>
</tr>
</table>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>D A D O S &nbsp;&nbsp;&nbsp; P E S S O A I S</b></td>
</tr>
</table>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Nome Completo:</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Endere�o:</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Bairro</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=25% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Cidade</b></td>
	<td class=campo width=15% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Estado</b></td>
	<td class=campo width=20% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>CEP</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Data de Nascimento</b></td>
	<td class=campo width=15% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Idade</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000" align="right">&nbsp;anos</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=27% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Telefone residencial</b></td>
	<td class=campo width=26% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>comercial/recados</b></td>
	<td class=campo width=26% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>celular</b></td>
	<td class=campo width=21% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Estado Civil</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td colspan=4 class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Sexo</b></td>
	<td class=campo width=80% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Nome do marido ou esposa</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="" align="left" width=16>Masculino</td>
	<td class=campo style="" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000" width=16>Feminino</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000" nowrap><b>Tem fihos?</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000" nowrap><b>Quantos?</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000" nowrap><b>Quantos menores de 14 anos?</b></td>
	<td class=campo width=70% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>E-mail</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Nome da M�e</b></td>
	<td class=campo width=50% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Nome do Pai</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=25% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Carteira Profissional</b></td>
	<td class=campo width=10% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>S�rie</b></td>
	<td class=campo width=15% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Data Emiss�o</b></td>
	<td class=campo width=30% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>T�tulo de Eleitor</b></td>
	<td class=campo width=10% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Zona</b></td>
	<td class=campo width=10% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Se��o</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=25% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>CPF n�</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>N� PIS / PASEP</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>R.G. (Identidade)</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Reservista N�</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>
<!--
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td colspan=4 class=campo style="border-top:1px solid #000000;border-right:1px solid #000000;border-left:1px solid #000000"><b>Se estrangeiro, qual o tipo de visto?</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Tempo de resid�ncia no pa�s?</b></td>
	<td colspan=4 class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>� deficiente f�sico?</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Que tipo de defici�ncia?</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="" align="left">at� ____/200___</td>
	<td class=campo style="" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000">Permanente</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="" align="left" width=16>Sim</td>
	<td class=campo style="" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000" width=16>N�o</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>
-->
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td colspan=4 class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000;border-left:1px solid #000000"><b>Tipo de visto? (se estrangeiro)</b></td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Tempo no pa�s?</b></td>
	<td colspan=4 nowrap class="campor" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Portador de defici�ncia?</b></td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Tipo?</b></td>
	<td colspan=4 class="campor" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Fumante?</b></td>
	<td colspan=4 class="campor" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Tatuagens?</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;" valign="center" width=10 ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo style="" align="left" nowrap >at� ____/20___</td>
	<td class=campo style="" width=10><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000">Permanente</td>

	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>

	<td class=campo style="border-left:1px solid #000000;" width=10 valign="center" ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo style="" align="left" width=16>Sim</td>
	<td class=campo style="" width=10><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000" >N�o</td>

	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>

	<td class=campo style="border-left:1px solid #000000;" width=10 valign="center" ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo style="" align="left" >Sim</td>
	<td class=campo style="" width=10 ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000" >N�o</td>

	<td class=campo style="border-left:1px solid #000000;" width=10 valign="center" ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo style="" align="left" >Sim</td>
	<td class=campo style="" width=10><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000">N�o</td>
	
</tr>
</table>


<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>C O N D I � � E S &nbsp;&nbsp;&nbsp; P A R A &nbsp;&nbsp;&nbsp; A D M I S S � O</b></td>
</tr>
</table>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Cargo</b></td>
	<td class=campo width=50% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Curso</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=campo colspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Disciplina(s)</b></td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;�</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;�</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;">&nbsp;�</td></tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=40% rowspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Tem parentes ou conhecidos<br>nesta Institui��o de Ensino?</b></td>
	<td class=campo valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Quem?</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>E D U C A � � O</b></td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=41% valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Escolaridade/Titula��o</b></td>
	<td class=campo width=19% rowspan=2 valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Estuda<br>atualmente?</b></td>
	<td class=campo width=40% valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>O que/onde?</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=4 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Cursos Gradua��o/Especializa��o/Etc.</b></td>
</tr>
<tr><td class=campo width=20% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Tipo</td>
	<td class=campo width=35% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Nome</td>
	<td class=campo width=35% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Institui��o</td>
	<td class=campo width=10% style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">Conclus�o</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<DIV style="page-break-after:always"></DIV>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>S I T U A � � O &nbsp;&nbsp;&nbsp; E C O N � M I C A</b></td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=15% style="border-top:1px solid #000000;border-right:1px solid #000000;border-left:1px solid #000000"><b>Banco</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Ag�ncia</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Nome da Ag�ncia</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Conta Corrente</b></td>
</tr>
<tr>
	<td class=campo style="border-right:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000;">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000;">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000;">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=5 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Ve�culos que utiliza (para emiss�o do crach� de estacionamento)</b></td>
</tr>
<tr><td class=campo width=25% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Marca</td>
	<td class=campo width=30% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Modelo</td>
	<td class=campo width=15% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Ano</td>
	<td class=campo width=15% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Cor</td>
	<td class=campo width=15% style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">Placa</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>E M P R E G O S &nbsp;&nbsp;&nbsp; A N T E R I O R E S</b></td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=3 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Relacione seus empregos anteriores para efeito de contagem de tempo de servi�o</b></td>
	<td class=campo colspan=3 valign=top style="border-top:1px solid #000000;border-right:1px solid #000000">N�o Preencher</td>
</tr>
<tr><td class=campo width=57% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Nome da Empresa/Institui��o</td>
	<td class=campo width=15% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Data Entrada</td>
	<td class=campo width=15% style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">Data Sa�da</td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000">A</td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000">M</td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000">D</td>
</tr>
<%for a=1 to 10%>
<tr><td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
</tr>
<%next%>
<tr><td class=campo colspan=3    style="border-top:2px solid #000000;border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Tempo de Servi�o (estimado)</td>
	<td class=campo align="center" style="border-top:2px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
	<td class=campo align="center" style="border-top:2px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
	<td class=campo align="center" style="border-top:2px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
</tr>
<tr><td class=campo colspan=3    style="border-top:2px solid #000000;border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Tempo restante para aposentadoria (estimado)</td>
	<td class=campo align="center" style="border-top:2px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
	<td class=campo align="center" style="border-top:2px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
	<td class=campo align="center" style="border-top:2px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
</tr>
</table>

<br>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Nome reduzido para crach� de identifica��o (at� 9 caracteres)</b></td>
</tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=campo colspan=2 style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	As declara��es aqui prestadas ser�o guardadas na mais restrita confian�a e fica subordinada � veracidade delas, 
	qualquer entendimento entre a Institui��o de Ensino e o candidato, <b>que se responsabiliza nos termos da lei pelas informa��es aqui declaradas.</td>
</tr>
<tr><td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Local e Data</b></td>
	<td class=campo width=50% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Assinatura</b></td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
</tr>
</table>


<br>
<br>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top:1px solid #000000">Recursos Humanos&nbsp;</td>
	<td class=campo style="border-top:1px solid #000000" align="right">&nbsp;Form. 10/2007</td>
</tr>
</table>
<%
%>
</body>
</html>