<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")="N" or session("a72")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Solicitação de Estágio</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<%
%>
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
	<tr>
		<td class=campo valign=top align="center" width=260><img src="../images/logo_centro_universitario_unifieo_big.gif" border="0" alt="" width="225">
		<p style="font-family:'Monotype Corsiva';font-size:14pt;margin-top:0;margin-bottom:0"><b>Tradição, qualidade, seriedade</p>
		<p style="font-family:'Monotype Corsiva';font-size:12pt;margin-top:0;margin-bottom:0">Desde 1967</p>
		</td>
		<td width=272 rowspan=2></td>
		<td class="campop" rowspan=2 width=118 height=157 style="border:1px solid #000000" align="center" valign="center">
		Foto<br>3 x 4<br>obrigatória</td>
	</tr>
	<tr>
		<td align="center"><p style="font-family:'Century Gothic';font-size:12pt;margin-top:0;margin-bottom:0"><b>Proposta de prestação de serviços</b></td>
	</tr>
</table>
<br>

<table border="0" cellpadding="1" cellspacing="4" width="650" style="border-collapse: collapse">
	<tr>
		<td class="campop" valign=top colspan=7><b>Vagas Pretendidas:</td></tr>
	<tr>
		<td class=campo width=30><b>.</td>
		<td class=campo width=15>&nbsp;</td>	
		<td class=campo>&nbsp;</td>	
		<td class=campo width=15><img src="../images/bola.gif" width=16 border="0"></td>	
		<td class=campo>7h às 13h</td>	
		<td class=campo width=300><b>.</td>
	</tr>
	<tr>
		<td class=campo><b>.</td>
		<td class=campo width=15>&nbsp;</td>	
		<td class=campo>&nbsp;</td>	
		<td class=campo width=15><img src="../images/bola.gif" width=16 border="0"></td>	
		<td class=campo>13h às 19h</td>	
		<td class=campo><b>.</td>
	</tr>
	<tr>
		<td class=campo><b>.</td>
		<td class=campo width=15>&nbsp;</td>	
		<td class=campo>&nbsp;</td>	
		<td class=campo width=15><img src="../images/bola.gif" width=16 border="0"></td>	
		<td class=campo>16h às 22h</td>	
		<td class=campo><b>.</td>
	</tr>
</table>
<br>

<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
	<tr><td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>Nome Completo:</b></td></tr>
	<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td></tr>
</table>
<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
	<tr><td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>Endereço:</b></td>
	<td class=campo width=29% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Bairro</b></td></tr>
	<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td>	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
	<tr><td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>Cidade</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Estado</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>CEP</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Data de Nascimento</b></td></tr>
	<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td></tr>
</table>
<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo width=27% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>Telefone residencial</b></td>
	<td class=campo width=26% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>comercial</b></td>
	<td class=campo width=26% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>celular</b></td>
	<td class=campo width=21% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Estado Civil</b></td>
	</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td>
	</tr>
</table>
<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
	<tr><td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>E-mail:</b></td></tr>
	<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
	<tr><td colspan=5 class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>Consulta seu e-mail frequentemente?</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>RG nº</b></td>
	<td class=campo width=35% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>CPF nº</b></td></tr>
	<tr><td class=campo style="border-left:1px solid #000000;" valign="center" width=16>
	<img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="" align="left" width=16>Sim</td>
	<td class=campo style="" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo align="left" style="" width=16>Não</td>

	<td class=campo style="border-right:1px solid #000000" width=46>
	&nbsp;&nbsp;&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
	<tr><td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>Filiação:</b></td></tr>
	<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo colspan=5 style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>É estudante do UNIFIEO?</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Caso não seja, qual a Instituição em que você estuda?</b></td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-bottom: 1px solid #000000" valign="center" width=16>
	<img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-bottom: 1px solid #000000" align="left" width=16>Sim</td>
	<td class=campo style="border-bottom: 1px solid #000000" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo align="left" style="border-bottom: 1px solid #000000" width=16>Não</td>
	<td class=campo style="border-right:1px solid #000000;border-bottom: 1px solid #000000" width=46>
	&nbsp;&nbsp;&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000;border-bottom: 1px solid #000000">
	&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="4" width="650" style="border-collapse: collapse">
	<tr>
		<td class=campo valign=top colspan=6><b>Curso:</td></tr>
	<tr>
		<td class=campo width=15><img src="../images/bola.gif" width=16 border="0"></td>	
		<td class=campo width=201>Publicidade e Propaganda</td>	
		<td class=campo width=15><img src="../images/bola.gif" width=16 border="0"></td>	
		<td class=campo width=201>Marketing</td>	
		<td class=campo width=15><img src="../images/bola.gif" width=16 border="0"></td>	
		<td class=campo width=201>Secretariado Executivo</td>	
	</tr>
	<tr>
		<td class=campo width=15><img src="../images/bola.gif" width=16 border="0"></td>	
		<td class=campo>Administração de Empresas</td>	
		<td class=campo width=15></td>	
		<td class=campo></td>	
		<td class=campo width=15></td>	
		<td class=campo></td>	
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=25% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>Prontuário</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Ano/Semestre</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Bloco</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Sala</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">
	&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo colspan=6 style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>Período</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Horário de 2ª a 6ª-feira</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">
	<b>Horário aos sábados</b></td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-bottom: 1px solid #000000" valign="center" width=16>
	<img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-bottom: 1px solid #000000" align="left">matutino</td>
	<td class=campo style="border-bottom: 1px solid #000000" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo align="left" style="border-bottom: 1px solid #000000">vespertino</td>
	<td class=campo style="border-bottom: 1px solid #000000" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000;border-bottom: 1px solid #000000">noturno</td>
	
	<td class=campo style="border-right:1px solid #000000;border-bottom: 1px solid #000000">
	das __________ às __________&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000;border-bottom: 1px solid #000000">
	das __________ às __________&nbsp;</td>
</tr>
</table>

<DIV style="page-break-after:always"></DIV>
<br>
<br>
<%pading=8%>
<table border="0" cellpadding="<%=pading%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>Possui veículo?</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Sim</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">Não</td>
</tr>	
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>Tem disponibilidade para trabalhar durante o período de férias?</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Sim</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">Não</td>
</tr>	
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>Tem disponibilidade para trabalhar aos sábados, domingos e feriados?</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Sim</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">Não</td>
</tr>	
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>Tem conhecimentos em informática?</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Sim</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">Não</td>
</tr>	
</table>
<table border="0" cellpadding="<%=pading%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>Em quais programas?</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Word</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Excel</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">PowerPoint</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Access</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Internet</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">Outro</td>
</tr>	
</table>
<table border="0" cellpadding="<%=pading%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>Sua digitação é:</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">lenta</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">moderada</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">rápida</td>
</tr>	
</table>
<table border="0" cellpadding="<%=pading%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>Uniforme:</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">pequeno</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">médio</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">grande</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">extra grande</td>
</tr>	
</table>
<table border="0" cellpadding="<%=pading%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>É fumante?</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Sim</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">Não</td>
</tr>	
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>Tem tatuagens?</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Sim</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">Não</td>
</tr>	
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;">
	<b>Já trabalhou em outros vestibulares da FIEO?</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000" align="left">Sim</td>
	<td class=campo style="border-top:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;" align="left">Não</td>
</tr>	
<tr>
	<td class=campo colspan=5 style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000;">
	<b>Quando?</td>
</tr>	
<tr>
	<td class=campo style="border-top:1px solid #000000;border-left:1px solid #000000;border-bottom:1px solid #000000;">
	<b>Caso seja estudante de Administração de Empresas:<br>Deseja se candidatar à vaga de Supervisor?</td>
	<td class=campo style="border-top:1px solid #000000;border-bottom:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-bottom:1px solid #000000;" align="left">Sim</td>
	<td class=campo style="border-top:1px solid #000000;border-bottom:1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="border-top: 1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000;" align="left">Não</td>
</tr>	
</table>

<br>
	
<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo><b>Experiências Anteriores:</td>
</tr>
</table>
<table border="1" bordercolor=#000000 cellpadding="8" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo rowspan=3 width=10 style="border-bottom: 2 solid #000000;"><b>1</td>
	<td class=campo><b>Empresa:</td>
</tr>
<tr>
	<td class=campo><b>Período:</td>
</tr>
<tr>
	<td class=campo style="border-bottom: 2 solid #000000;"><b>Função:</td>
</tr>
<tr>
	<td class=campo rowspan=3 width=10 style="border-bottom: 2 solid #000000;"><b>2</td>
	<td class=campo><b>Empresa:</td>
</tr>
<tr>
	<td class=campo><b>Período:</td>
</tr>
<tr>
	<td class=campo style="border-bottom: 2 solid #000000;"><b>Função:</td>
</tr>
<tr>
	<td class=campo rowspan=3 width=10><b>3</td>
	<td class=campo><b>Empresa:</td>
</tr>
<tr>
	<td class=campo><b>Período:</td>
</tr>
<tr>
	<td class=campo><b>Função:</td>
</tr>
</table>
<br>
<br>
<br>
<br>
<br>
<br>
<p style="margin-top:0;margin-bottom:0">Osasco, ______ de __________________________ de <%=year(now)%></p>
<br>
<br>
<br>
<br>
<p style="margin-top:0;margin-bottom:0">Assinatura:</p>	

<%
%>
</body>
</html>