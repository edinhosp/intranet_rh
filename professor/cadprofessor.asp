<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a38")="N" or session("a38")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastro Professor Convidado</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<%
espacamento=5
%>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop" width=240 valign=top align="left"><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td>
	<td class="campop" width=410 valign="center" align="center"><p style="font-family:'Century Gothic';font-size:16pt;margin-top:0;margin-bottom:0"><b>Cadastro de Professor Convidado</b></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop" align="center" valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">
	<b>D A D O S &nbsp;&nbsp;&nbsp; P E S S O A I S</b></td>
</tr>
</table>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Nome Completo:</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Apelido/Nome informal</b></td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Data de Nascimento</b></td>
	<td colspan=4 class=campo style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Sexo</b></td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Nacionalidade</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-left: 1px solid #000000;" valign="center" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo style="" align="left" width=16>Masculino</td>
	<td class=campo style="" width=16><img src="../images/bola.gif" width=16 border="0"></td>
	<td class=campo align="left" style="border-right: 1px solid #000000" width=16>Feminino</td>
	<td class=campo style="border-right: 1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Naturalidade</b></td>
	<td class=campo width=15% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Estado Natal</b></td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Estado Civil</b></td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Grau de Instrução</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Nome da Mãe</b></td>
	<td class=campo width=50% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Nome do Pai</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Email</b></td>
<!--	<td class=campo width=30% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Titulação</b></td> -->
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
<!--	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td> -->
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop" align="center" valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">
	<b>E N D E R E Ç O &nbsp;&nbsp;&nbsp; P R I N C I P A L</b></td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Endereço</b></td>
	<td class=campo width=10% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Número</b></td>
	<td class=campo width=25% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Complemento</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Bairro</b></td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Cidade</b></td>
	<td class=campo width=10% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Estado</b></td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>CEP</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=27% style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Telefone residencial</b></td>
	<td class=campo width=26% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>comercial/recados</b></td>
	<td class=campo width=26% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>celular</b></td>
	<td class=campo width=21% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>fax</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop" align="center" valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">
	<b>D O C U M E N T O S</b></td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=25% style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Carteira Profissional</b></td>
	<td class=campo width=10% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Série</b></td>
	<td class=campo width=15% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Data Emissão</b></td>
	<td class=campo width=30% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Título de Eleitor</b></td>
	<td class=campo width=10% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Zona</b></td>
	<td class=campo width=10% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Seção</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo width=25% style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>CPF nº</b></td>
	<td class=campo width=25% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Nº PIS / PASEP</b></td>
	<td class=campo width=25% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>R.G. (Identidade)</b></td>
	<td class=campo width=25% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Reservista Nº</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000" align="right">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop" align="center" valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">
	<b>F O R M A Ç Ã O &nbsp;&nbsp;&nbsp; A C A D Ê M I C A</b></td>
</tr>
<tr>
	<td class="campop" align="left" style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">
	Tipos de formação acadêmcia: Graduação/Especialização/Mestrado/Doutorado</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Tipo</b></td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Curso</b></td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Instituição</	b></td>
	<td class=campo width=15% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Data conclusão</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=campo style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=campo style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=campo style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo colspan=2 valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Como entrou em contato com esta Instituição?</b></td>
</tr>
<tr>
	<td class=campo colspan=2 style="border-left: 1px solid #000000;border-right: 1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=campo colspan=2 style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000">
	As declarações aqui prestadas serão guardadas na mais restrita confiança e fica subordinada à veracidade delas, 
	qualquer entendimento entre a Instituição de Ensino e o candidato.</td>
</tr>
<tr>
	<td class=campo valign=top style="border-top: 1px solid #000000;border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Local e Data</b></td>
	<td class=campo width=50% style="border-top: 1px solid #000000;border-right: 1px solid #000000"><b>Assinatura</b></td>
</tr>
<tr>
	<td class=campo style="border-left: 1px solid #000000;border-right: 1px solid #000000;border-bottom: 1px solid #000000">&nbsp;</td>
	<td class=campo style="border-right: 1px solid #000000;border-bottom: 1px solid #000000">&nbsp;</td>
</tr>
</table>

<br>
<br>
<br>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top: 1px solid #000000">Recursos Humanos&nbsp;</td>
	<td class=campo style="border-top: 1px solid #000000" align="right">&nbsp;Form. 04/2004</td>
</tr>
</table>

<%
%>
</body>
</html>