<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a43")="N" or session("a43")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>RDT - Retifica��o de Dados do Trabalhador</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"><!--
function nome1() {	form.chapa.value=form.nome.value; }
function chapa1() {	form.nome.value=form.chapa.value; }
--></script>
</head>
<body style="margin-left:20px">
<%
espacamento=5
%>
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top><img src="../images/rdt_caixa.gif" height="33" border="0"></td>
	<td class=campo valign=top><img src="../images/rdt_prev2.gif" height="42" border="0"></td>
	<td class=campo valign=top><img src="../images/rdt_mtb.jpg"   height="36" border="0"></td>
	<td class="campop" valign=top width=195><b>R D T - Retifica��o de Dados<br>do Trabalhador FGTS/INSS<br>Modelo 4</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="650" style="border-collapse: collapse"><tr><td width=445>
<!-- 1 identifica��o -->
<table border="0" cellpadding="1" cellspacing="0" width="445" style="border-collapse: collapse">
<tr><td class="campor"><b>1 - IDENTIFICA��O DO EMPREGADOR/CONTRIBUINTE (preenchimento obrigat�rio)</td></tr>
<tr><td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Raz�o Social/Nome</td></tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="FUNDA��O INSTITUTO DE ENSINO PARA OSASCO"></td></tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="445" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Codigo empregador/contribuinte (empresas com FGTS)</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF conta</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ/CEI do empregador/contribuinte</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="35" value="06951101389480"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="3" value="SP"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="25" value="73.066.166/0003-92"></td>
</tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" width="445" style="border-collapse: collapse">
<tr><td class="campor" height=5></td></tr>
<tr><td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Pessoa para Contato/DDD/Telefone</td></tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="ROGERIO / 011 / 3651-9905"></td></tr>
<tr><td class="campor" height=5></td></tr>
</table>

<!-- 2 identifica��o -->
<table border="0" cellpadding="1" cellspacing="0" width="445" style="border-collapse: collapse">
<tr><td class="campor"><b>2 - IDENTIFICA��O DO TRABALHADOR (preenchimento obrigat�rio)</td></tr>
<tr><td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome do Trabalhador</td></tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value=""></td></tr>
<tr><td class="campor" height=5></td></tr>
</table>
</td>
<td valign=top align="right">
	<table border="0" cellpadding="1" cellspacing="0" width="197" style="border-collapse: collapse">
	<tr><td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Carimbo CIEF</td></tr>
	<tr><td class=campo height=142 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
	</table>
</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;N� do PIS/PASEP/inscri��o contribuinte individual</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Admiss�o</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;C�digo do Trabalhador (categorias com FGTS)</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Categoria</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="35" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="12" value="  /  /    "></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="35" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="01"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>

<!-- 3 dados cadastrais -->
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class="campor" colspan=2><b>3 - DADOS CADASTRAIS (preencher somente os campos a serem alterados)</td></tr>
<tr><td class="campor" width=430 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome do Trabalhador</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;N�PIS/PASEP/inscri��o contribuinte individual</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="30" value=""></td>
</tr>
<tr><td class="campor" colspan=2 height=5></td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campor" width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Admiss�o</td>
	<td class="campor" width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de op��o</td>
	<td class="campor" width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de retroa��o</td>
	<td class="campor" width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Nascimento</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Movimenta��o informada</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Movimenta��o correta</td>
</tr>
<tr>
	<td class="campor" width=100 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data</td>
	<td class="campor" width=50  style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;C�digo</td>
	<td class="campor" width=100 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data</td>
	<td class="campor" width=50  style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;C�digo</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" width=650 style="border-collapse: collapse">
<tr>
	<td class="campor" width=50 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Categoria</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Matr�cula</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;N� CTPS</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;S�rie</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Unidade de Trabalho</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="3" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="7" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="3" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value=""></td>
</tr>
<tr><td class="campor" colspan=6 height=5></td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" width=650 style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Endere�o (logradouro, n�mero, andar, apartamento, etc.)</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="110" value=""></td>
</tr>
<tr><td class="campor" colspan=1 height=5></td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" width=650 style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Bairro/Distrito</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Munic�pio</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CEP</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="30" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="30" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="3" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value=""></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>

<!-- 4 dados a retificar -->
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class="campor"><b>4 - DADOS A RETIFICAR POR GFIP/GRFP/GRFC (anexar GFIP/GRFP/GRFC incorreta e nova(s) GFIP/GRFP/GRFC, conforme o caso.)</td></tr>
<tr><td class="campor"><b>IDENTIFICA��O DO RECOLHIMENTO / DECLARA��O (preenchimento obrigat�rio)</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campor" rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Banco</td>
	<td class="campor" rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Ag�ncia</td>
	<td class="campor" rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Compet�ncia</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;C�digo de</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ/CEI tomador de servi�o</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;M�s</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Ano</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;recolhimento</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;/obra constr.civil informado</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="3" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="4" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="4" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value=""></td>
</tr>
<tr><td class="campor" colspan=10 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class="campor" colspan=5><b>RETIFICA��O DOS DADOS (preencher somente os dados a serem alterados)</td></tr>
<tr>
	<td class="campor" width=447 colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ/CEI do empregador/contribuinte</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" width=200 colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;FPAS</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;INFORMADO</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CORRETO</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;INFORMADO</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CORRETO</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo width=110 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value=""></td>
	<td class=campo width=110 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value=""></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campor" rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Valor descontado do segurado</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Valor base de c�lculo 13�sal�rio da Previd�ncia Social</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Valor base de c�lculo da Previd�ncia Social</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;- Referente compet�ncia de movimento</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;- acidente de trabalho/servi�o militar obrigat�rio</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value=""></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campor" rowspan=2 valign=middle align="center" style="">&nbsp;Correto</td>
	<td class="campor" rowspan=2 valign=middle align="center" style="">&nbsp;&nbsp;<img src="../images/setanext1.gif" width="12" height="12" border="0" alt=""></td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Raz�o Social do tomador de servi�o/obra de constru��o civil</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ/CEI do tomador de servi�o/obra de constru��o civil</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="60" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="30" value=""></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>

<!-- 5 dados a retificar -->
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class="campor" colspan=11><b>5 - DADOS A RETIFICAR POR PER�ODO</td></tr>
<tr><td class="campor" colspan=6><b>IDENTIFICA��O DO PER�ODO</td>
	<td class="campor" colspan=5><b>RETIFICA��O DOS DADOS</td></tr>
<tr>
	<td class="campor" colspan=5 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Compet�ncia (preencimento obrigat�rio)</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CBO</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;C�digo de ocorr�ncia</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;M�s</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Ano</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;at�</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;M�s</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Ano</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;INFORMADO</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CORRETO</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;INFORMADO</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CORRETO</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="7" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="7" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="7" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="7" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="7" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="7" value=""></td>
</tr>
<tr><td class="campor" colspan=11 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=3><b>Poder�o ser exigidos outros documentos, caso a CAIXA julgue necess�rio.</td></tr>
<tr>
	<td width=48% class="campop" style="">
		&nbsp;<input class=form_input type="text" name="razao" size="50" value=""></td>
	<td class="campop" style="">&nbsp;</td>
	<td width=48% class="campop" style="">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
</tr>
<tr>
	<td class=campo style="border-top:1px solid #000000">Local e Data</td>
	<td class="campop" style="">&nbsp;</td>
	<td class=campo style="border-top:1px solid #000000">Carimbo e assinatura do respons�vel<br>
	<input type="text" class=form_input size="60" value="ROGERIO MATEUS DOS SANTOS ARAUJO-CPF 185.420.058-56"></td>
<!--RG 27.831.325-5 -->
</tr>
<tr><td class="campor" colspan=3 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=3><b>PARA USO DA CAIXA</td></tr>
<tr><td class=campo colspan=3>Declaro que os documentos apresentados comprovam as altera��es solicitadas.</td></tr>
<tr><td class=campo colspan=1></td>
	<td class="campor" colspan=2>Novas GFIP/GRFP/GRFC Anexadas</td></tr>
<tr>
	<td class=campo width=400 style="border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo width=30 style="border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class="campor" width=220 style="">0 - Sim<br>1 - N�o</td>
</tr>
<tr><td class="campor" colspan=3>Assinaturo/carimbo do respons�vel pela confer�ncia</td></tr>
</table>
<%
%>
</body>
</html>