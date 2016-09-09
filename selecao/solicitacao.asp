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
<title>Solicitação de Emprego</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<%
espacamento=5
%>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop" width=240 rowspan=3 valign=top align="left"><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td>
	<td class="campop" width=325 height=71><p style="font-family:'Century Gothic';font-size:18pt;margin-top:0;margin-bottom:0"><b>Solicitação de Emprego</b></td>
	<td class=campo width=15 valign=bottom></td>	
	<td class=campo width=70 valign=bottom></td>	
</tr>
<tr>
	<td class=campo height=20>&nbsp;</td>	
	<td class=campo width=15></td>	
	<td class=campo width=70></td>	
</tr>
</table>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
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
	<td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Endereço:</b></td>
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
	<td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Nome da Mãe</b></td>
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
	<td class=campo width=10% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Série</b></td>
	<td class=campo width=15% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Data Emissão</b></td>
	<td class=campo width=30% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Título de Eleitor</b></td>
	<td class=campo width=10% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Zona</b></td>
	<td class=campo width=10% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Seção</b></td>
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
	<td class=campo width=25% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>CPF nº</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Nº PIS / PASEP</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>R.G. (Identidade)</b></td>
	<td class=campo width=5% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Estado</b></td>
	<td class=campo width=20% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Data Emissão</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td colspan=4 class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000;border-left:1px solid #000000"><b>Tipo de visto? (se estrangeiro)</b></td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Tempo no país?</b></td>
	<td colspan=4 nowrap class="campor" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Portador de deficiência?</b></td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Tipo?</b></td>
	<td colspan=4 class="campor" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Fumante?</b></td>
	<td colspan=4 class="campor" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Tatuagens?</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;" valign="center" width=10 ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo style="" align="left" nowrap >até ____/20___</td>
	<td class=campo style="" width=10><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000">Permanente</td>

	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>

	<td class=campo style="border-left:1px solid #000000;" width=10 valign="center" ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo style="" align="left" width=16>Sim</td>
	<td class=campo style="" width=10><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000" >Não</td>

	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>

	<td class=campo style="border-left:1px solid #000000;" width=10 valign="center" ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo style="" align="left" >Sim</td>
	<td class=campo style="" width=10 ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000" >Não</td>

	<td class=campo style="border-left:1px solid #000000;" width=10 valign="center" ><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo style="" align="left" >Sim</td>
	<td class=campo style="" width=10><img src="../images/bullet2.gif" border="0"></td>
	<td class=campo align="left" style="border-right:1px solid #000000">Não</td>
	
</tr>
</table>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>C O N D I Ç Õ E S &nbsp;&nbsp;&nbsp; P A R A &nbsp;&nbsp;&nbsp; A D M I S S Ã O</b></td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Que cargo pretende?</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Pretensão salarial</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=33% rowspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Submete-se a um prazo de<br>3 meses para experiência?</b></td>
	<td class=campo width=27% rowspan=2 valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Tem preferência por<br>horário de trabalho?</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Qual?</b></td>
</tr>
<tr>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=35% rowspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Tem parentes ou conhecidos<br>nesta Instituição de Ensino?</b></td>
	<td class=campo valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Quem?</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>E D U C A Ç Ã O</b></td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=41% valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Escolaridade</b></td>
	<td class=campo width=19% rowspan=2 valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Estuda<br>atualmente?</b></td>
	<td class=campo width=40% valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>O que/onde?</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=28% rowspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Possui cursos<br>extras curriculares?</b></td>
	<td class=campo valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Quais?</b></td>
</tr>
<tr>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>S I T U A Ç Ã O &nbsp;&nbsp;&nbsp; E C O N Ô M I C A</b></td>
</tr>
</table>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=21% rowspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Possui alguma<br>propriedade?</b></td>
	<td class=campo width=16% rowspan=2 valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Está livre<br>de ônus?</b></td>
	<td class=campo width=18% rowspan=2 valign=top style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Tem<br>automóvel?</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Qual?</b></td>
	<td class=campo width=17% rowspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000;"><b>Tem seguro<br>de vida?</b></td>
</tr>
<tr>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=15% style="border-top:1px solid #000000;border-right:1px solid #000000;border-left:1px solid #000000"><b>Banco</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Agência</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Nome da Agência</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Conta Corrente</b></td>
</tr>
<tr>
	<td class=campo style="border-bottom:1px solid #000000;border-right:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-bottom:1px solid #000000;border-right:1px solid #000000;">&nbsp;</td>
	<td class=campo style="border-bottom:1px solid #000000;border-right:1px solid #000000;">&nbsp;</td>
	<td class=campo style="border-bottom:1px solid #000000;border-right:1px solid #000000;">&nbsp;</td>
</tr>
</table>

<DIV style="page-break-after:always"></DIV>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>E M P R E G O S &nbsp;&nbsp;&nbsp; A N T E R I O R E S</b></td>
</tr>
</table>
<%for a=1 to 2%>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=3% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b><%=a%></b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Nome da Empresa</b></td>
	<td class=campo width=30% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Nome do Superior imediato</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=3% style="border-left:1px solid #000000;border-right:1px solid #000000"></td>
	<td class=campo width=20% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Cidade</b></td>
	<td class=campo width=25% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Telefone</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Cargo</b></td>
	<td class=campo width=18% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Salário</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=3% style="border-left:1px solid #000000;border-right:1px solid #000000"></td>
	<td class=campo width=19% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Data Entrada</b></td>
	<td class=campo width=19% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Data Saída</b></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Motivo da Saída</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>
<%next%>

<!-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -->
<!--
<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=fundop align="center" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b>R E F E R Ê N C I A S</b></td>
</tr>
</table>
-->

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=3 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Relacione seus empregos anteriores para efeito de contagem de tempo de serviço</b></td>
	<td class=campo colspan=3 valign=top style="border-top:1px solid #000000;border-right:1px solid #000000">Não Preencher</td>
</tr>
<tr><td class=campo width=57% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Nome da Empresa/Instituição</td>
	<td class=campo width=15% style="border-left:1px solid #000000;border-bottom:1px solid #000000">Data Entrada</td>
	<td class=campo width=15% style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">Data Saída</td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000">A</td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000">M</td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000">D</td>
</tr>
<%for a=3 to 9%>
<tr><td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000"><%=a%>-&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
	<td class=campo align="center" style="border-right:1px solid #000000;border-bottom:1px solid #000000"></td>
</tr>
<%next%>
<tr><td class=campo colspan=3    style="border-top:2px solid #000000;border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Tempo de Serviço (estimado)</td>
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
<tr>
	<td class=campo colspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Como entrou em contato com esta Instituição?</b></td>
</tr>
<tr>
	<td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=campo colspan=2 style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	As declarações aqui prestadas serão guardadas na mais restrita confiança e fica subordinada à veracidade delas, 
	qualquer entendimento entre a Instituição de Ensino e o candidato, <b>que se responsabiliza nos termos da lei pelas informações aqui declaradas.</td>
</tr>
<tr>
	<td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Local e Data</b></td>
	<td class=campo width=50% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Assinatura</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
</tr>
</table>


<hr>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Entrevistado por</b></td>
	<td class=campo width=50% style="border-top:1px solid #000000;border-right:1px solid #000000"><b>Local/Data</b></td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=campo colspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Observações</b></td>
</tr>
<tr>
	<td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
</tr>
</table>

<br>
<br>
<br>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top:1px solid #000000">Recursos Humanos&nbsp;</td>
	<td class=campo style="border-top:1px solid #000000" align="right">&nbsp;Form. 04/2004</td>
</tr>
</table>

<%
%>
</body>
</html>