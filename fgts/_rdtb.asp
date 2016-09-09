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
<title>RDT - Retificação de Dados do Trabalhador</title>
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
	<td class="campop" valign=top width=195><b>R D T - Retificação de Dados<br>do Trabalhador FGTS/INSS<br>Modelo 4</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="650" style="border-collapse: collapse"><tr><td width=445>
<!-- 1 identificação -->
<table border="0" cellpadding="1" cellspacing="0" width="445" style="border-collapse: collapse">
<tr><td class="campor"><b>1 - IDENTIFICAÇÃO DO EMPREGADOR/CONTRIBUINTE (preenchimento obrigatório)</td></tr>
<tr><td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Razão Social/Nome</td></tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO"></td></tr>
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

<!-- 2 identificação -->
<table border="0" cellpadding="1" cellspacing="0" width="445" style="border-collapse: collapse">
<tr><td class="campor"><b>2 - IDENTIFICAÇÃO DO TRABALHADOR (preenchimento obrigatório)</td></tr>
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
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº do PIS/PASEP/inscrição contribuinte individual</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Admissão</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Código do Trabalhador (categorias com FGTS)</td>
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
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;NºPIS/PASEP/inscrição contribuinte individual</td>
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
	<td class="campor" width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Admissão</td>
	<td class="campor" width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de opção</td>
	<td class="campor" width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de retroação</td>
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
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Movimentação informada</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Movimentação correta</td>
</tr>
<tr>
	<td class="campor" width=100 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data</td>
	<td class="campor" width=50  style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Código</td>
	<td class="campor" width=100 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data</td>
	<td class="campor" width=50  style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Código</td>
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
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Matrícula</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº CTPS</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Série</td>
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
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Endereço (logradouro, número, andar, apartamento, etc.)</td>
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
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Município</td>
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
<tr><td class="campor"><b>IDENTIFICAÇÃO DO RECOLHIMENTO / DECLARAÇÃO (preenchimento obrigatório)</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campor" rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Banco</td>
	<td class="campor" rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Agência</td>
	<td class="campor" rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Competência</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Código de</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ/CEI tomador de serviço</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Mês</td>
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
<tr><td class="campor" colspan=5><b>RETIFICAÇÃO DOS DADOS (preencher somente os dados a serem alterados)</td></tr>
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
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Valor base de cálculo 13ºsalário da Previdência Social</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Valor base de cálculo da Previdência Social</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;- Referente competência de movimento</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;- acidente de trabalho/serviço militar obrigatório</td>
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
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Razão Social do tomador de serviço/obra de construção civil</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ/CEI do tomador de serviço/obra de construção civil</td>
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
<tr><td class="campor" colspan=11><b>5 - DADOS A RETIFICAR POR PERÍODO</td></tr>
<tr><td class="campor" colspan=6><b>IDENTIFICAÇÃO DO PERÍODO</td>
	<td class="campor" colspan=5><b>RETIFICAÇÃO DOS DADOS</td></tr>
<tr>
	<td class="campor" colspan=5 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Competência (preencimento obrigatório)</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CBO</td>
	<td class="campor" width=3 rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Código de ocorrência</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Mês</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Ano</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;até</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Mês</td>
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
<tr><td class=campo colspan=3><b>Poderão ser exigidos outros documentos, caso a CAIXA julgue necessário.</td></tr>
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
	<td class=campo style="border-top:1px solid #000000">Carimbo e assinatura do responsável<br>
	<input type="text" class=form_input size="60" value="ROGERIO MATEUS DOS SANTOS ARAUJO-CPF 185.420.058-56"></td>
<!--RG 27.831.325-5 -->
</tr>
<tr><td class="campor" colspan=3 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=3><b>PARA USO DA CAIXA</td></tr>
<tr><td class=campo colspan=3>Declaro que os documentos apresentados comprovam as alterações solicitadas.</td></tr>
<tr><td class=campo colspan=1></td>
	<td class="campor" colspan=2>Novas GFIP/GRFP/GRFC Anexadas</td></tr>
<tr>
	<td class=campo width=400 style="border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo width=30 style="border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class="campor" width=220 style="">0 - Sim<br>1 - Não</td>
</tr>
<tr><td class="campor" colspan=3>Assinaturo/carimbo do responsável pela conferência</td></tr>
</table>
<%
%>
</body>
</html>