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
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sessao=session.sessionid
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

espacamento=5
if request.form="" then
sql="select p.chapa, p.nome from corporerm.dbo.pfunc p where p.chapa<'10000' and p.codtipo='N' order by p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="form" action="rdt.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleção de Funcionário para emissão de RDT</td>
</tr>
<tr>
	<td class=campo>Funcionário</td>
	<td class=campo><input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()"></td>
	<td class=campo>
		<select name="nome" class=a onchange="nome1()">
		<option value="0"> Selecione o funcionário</option>
<%
rs.movefirst
do while not rs.eof
%>
		<option value="<%=rs("chapa")%>"> <%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
%>
		</select>
	</td>
</tr>
<tr>
	<td class=campo colspan=3>&nbsp;
		<input type="submit" value="Visualizar" class=button name="B1">
	</td>
</tr>
</table>
</form>

<%
else
sql="select f.chapa, f.nome, f.pispasep, f.dataadmissao, f.contafgts, f.codcategoria, f.dtopcaofgts, p.dtnascimento, " & _
"p.carteiratrab, p.seriecarttrab, p.ufcarttrab, p.rua, p.numero, p.complemento, p.bairro, p.cidade, p.estado, p.cep, f.codsecao " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and f.chapa='" & request.form("chapa") & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
campus=left(rs("codsecao"),2)
if campus="01" then cnpj="73.063.166/0001-20"
if campus="03" then cnpj="73.063.166/0003-92" 
if campus="04" then cnpj="73.063.166/0004-73" 
categoria=numzero(rs("codcategoria"),2)
largura=1050 '650
larg2=800 '445
%>
<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top width=175><img src="../images/rdt_caixa.gif" height="33" border="0"></td>
	<td class="campop" valign=middle align="left"><b>R D T - Retificação de Dados do Trabalhador - FGTS</td>
	<td width=70>&nbsp;</td>
</tr>
<tr>
	<td colspan=2 width=980>&nbsp;</td>
	<td class=campo height=35 width=70 style="border-left: 1px solid;border-bottom: 1px solid;border-right: 1px solid" valign=top>Grau de sigilo<br>&nbsp;</td>
</tr>
<tr>
	<td class=campo colspan=3><b>Orientações de preenchimento são obtidas no "Manual de Orientações, Retificação de Dados, Transferência de Contas
	Vinculadas e Devolução de Valores Recolhidos a Maior", disponível no sítio da CAIXA na Internet > downloads > FGTS > extrato e retificação de dados.
	</td>
</tr>
</table>

<table><tr><td class="campor" height=5></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td width=<%=larg2%>>
<!-- 1 identificação -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=larg2%>" style="border-collapse: collapse">
<tr><td class=campo colspan=5><b>1 - Identificação do Empregador</b> (Preenchimento obrigatório. Informar dados do cadastro do FGTS)</td></tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Razão Social/Nome</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ/CEI do empregador</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Codigo do empregador</td>
	<td class=campo style="border-left:1px solid #000000;border-right:0px solid #000000">&nbsp;Base da conta</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="50" value="FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%=cnpj%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="3" value="SP"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="25" value="06951101389480"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:0px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="3" value="SP"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=larg2%>" style="border-collapse: collapse">
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Pessoa para Contato</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;DDD/telefone</td>
	<td class=campo style="border-left:1px solid #000000;border-right:0px solid #000000">&nbsp;Endereço eletrônico (e-mail)</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="35" value="ROGERIO"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="15" value="(011) 3651-9905"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:0px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="50" value="rogerio@unifieo.br"></td>
	</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<!-- 2 identificação -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=larg2%>" style="border-collapse: collapse">
<tr><td class=campo colspan=5><b>2 - Identificação do Trabalhador</b> (Preenchimento obrigatório. Informar dados do cadastro do FGTS, mesmo que incorretos)</td></tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome do Trabalhador</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº do PIS/PASEP</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Admissão</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Categoria</td>
	<td class=campo style="border-left:1px solid #000000;border-right:0px solid #000000">&nbsp;Código do Trabalhador</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="50" value="<%=rs("nome")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("pispasep")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="12" value="<%=rs("dataadmissao")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=categoria%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:0px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("contafgts")%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<!-- 3 dados cadastrais -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=larg2%>" style="border-collapse: collapse">
<tr><td class=campo colspan=6><b>3 - Dados Cadastrais a Retificar</b> (Preencher, somente, os campos a serem alterados)</td></tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome do Trabalhador</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº PIS/PASEP</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº CTPS</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Série</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
	<td class=campo style="border-left:1px solid #000000;border-right:0px solid #000000">&nbsp;Categoria</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="60" value="<%=rs("nome")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("pispasep")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="7" value="<%=rs("carteiratrab")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=rs("seriecarttrab")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="3" value="<%=rs("ufcarttrab")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:0px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=categoria%>"></td>
</tr>
<tr><td class="campor" colspan=6 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=campo width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Admissão</td>
	<td class=campo width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de opção</td>
	<td class=campo width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de retroação</td>
	<td class=campo width=130 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Nascimento</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("dataadmissao")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("dtopcaofgts")%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("dtnascimento")%>"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>

</td>
<td valign=top align="right">
	<table border="0" cellpadding="1" cellspacing="0" width="250" style="border-collapse: collapse">
	<tr><td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;PARA USO DA CAIXA<br>
	Protocolo de recepção e assinatura, sob carimbo,<br>do responsável pela conferência</td></tr>
	<tr><td class=campo height=200 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
	</table>
</td>
</tr>
</table>

<!-- fim da primeira parte -->

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr><td class=campo colspan=3><b>4 - Pedido de Exclusão de Movimentação Informada</b> (Preencher com o dado informado indevidamente. Aplicado somente para exclusão de informação prestada)</td></tr>
<tr>
	<td class=campo width=100 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data</td>
	<td class=campo width=50  style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Código</td>
	<td class=campo width="90%" style="">&nbsp;</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
	<td class=campo style="">&nbsp;</td>
</tr>
<tr><td class="campor" colspan=3 height=5></td></tr>
</table>

<!-- 4 dados a retificar -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class=campo><b>5 - Retificação da remuneração sem devolução do FGTS, no mesmo Empregador, na mesma Competência e entre contas do mesmo trabalhador ou entre trabalhadores diferentes.</td></tr>
<tr><td class="campor"><b>* Nas guias com recolhimento ao FGTS, as remunerações informadas no campo "PARA" devem ser limitadas aos valores discriminados no campo "DE"</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class="campor" colspan=5 height=15 style="border-left: 1px solid #000000"><b>De:</b> (Preencher com dados informados incorretamente na guia)</td>
	<td class="campor" colspan=6 height=15 style="border-left: 1px solid #000000;border-right: 1px solid #000000"><b>Para:</b> (Preencher com dados corretos para a guia)</td>
</tr>
<tr>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000" height=20>&nbsp;Nome do trabalhador</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº do PIS/PASEP</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Categoria</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de admissão</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Remuneração&nbsp;</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome do trabalhador</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº do PIS/PASEP</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Categoria</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de admissão</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Remuneração</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Competência</td>
</tr>
<%for a=1 to 3%>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000" height=20>
		<%=a%>-&nbsp;<input class=form_input7 type="text" name="razao" size="40" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input7 type="text" name="razao" size="11" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input7 type="text" name="razao" size="2" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input7 type="text" name="razao" size="12" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input7 type="text" name="razao" size="10" value="         ,"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=a%>-&nbsp;<input class=form_input7 type="text" name="razao" size="40" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input7 type="text" name="razao" size="11" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input7 type="text" name="razao" size="2" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input7 type="text" name="razao" size="12" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input7 type="text" name="razao" size="10" value="         ,"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input7 type="text" name="razao" size="8" value="   /"></td>
</tr>
<%next%>
<tr><td class="campor" colspan=10 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class=campo colspan=5 height=15 style=""><b>6 - Pedido de Unificação de Contas dos Trabalhadores em Multiplicidade</b></td>
	<td class="campor">&nbsp;</td>
	<td class=campo colspan=5 height=15 style=""><b>7 - Pedido de Atualização de Saque na Vigência do Contrato de Trabalho</b></td>
</tr>
<tr><td class=campo colspan=5 height=15 style="border-left: 1px solid #000000;border-right: 1px solid #000000">Código das contas vinculadas do trabalhador a serem unificadas</td>
	<td class="campor">&nbsp;</td>
	<td class=campo colspan=5 height=15 style="border-left: 1px solid #000000;border-right: 1px solid #000000">Código da conta vinculada a ser atualizada</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000" height=20>
		<input class=form_input type="text" name="razao" size="20" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000" height=20>
		<input class=form_input type="text" name="razao" size="20" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000" height=20>
		<input class=form_input type="text" name="razao" size="20" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000" height=20>
		<input class=form_input type="text" name="razao" size="20" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000" height=20>
		<input class=form_input type="text" name="razao" size="20" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000" height=20>
		<input class=form_input type="text" name="razao" size="20" value=""></td>
</tr>
<tr><td class="campor" colspan=11 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class="campor" colspan=3><b style="font-family:Arial;font-size:9px">Estou ciente de que se verificada, a qualquer tempo, a falsidade das informações constantes deste documento,
sujeitar-se-á o responsável às penalidades previstas na legislação civil e penal, sem prejuízo das ações administrativas cabíveis.</td></tr>
<tr>
	<td width=38% class="campop" style="" height=30>
		&nbsp;<input class=form_input type="text" name="razao" size="50" value="<%="OSASCO, " & day(now) & " DE " & ucase(monthname(month(now))) & " DE " & year(now)%>"></td>
	<td class="campop" style="">&nbsp;</td>
	<td width=58% class="campop" style="">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
</tr>
<tr>
	<td class=campo style="border-top:1px solid #000000" valign=top>Local/Data</td>
	<td class="campop" style="">&nbsp;</td>
	<td class=campo style="border-top:1px solid #000000">Identificação e assinatura do responsável pela empresa ou seu representante legal<br>
	NOME: <input type="text" class=form_input size="60" value="ROGERIO MATEUS DOS SANTOS ARAUJO"><br>
	CPF: <input type="text" class=form_input size="20" value="185.420.058-56">
	</td>
<!--RG 27.831.325-5-->
</tr>
<tr><td class="campor" colspan=3 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td colspan=3 align="center">
	<p style="font-size:9pt;margin-top:0px;margin-bottom:0px;"><b>Documento não aplicável ao Recolhimento Rescisório</b></p>
	<p style="font-size:8pt;margin-top:0px;margin-bottom:0px;"><b>SAC CAIXA:</b> 0800 726 0101 (informações, reclamações, sugestões e elogios)</p>
	<p style="font-size:7pt;margin-top:0px;margin-bottom:0px;"><b>Para pessoas com deficiência auditiva:</b> 0800 726 2492</p>
	<p style="font-size:7pt;margin-top:0px;margin-bottom:0px;"><b>Ouvidoria:</b> 0800 725 7474 (reclamações não solucionadas e denúncias)</p>
</td></tr>
<tr><td class="campor" colspan=1 align="left" valign=top width="33%">31.004 v014 micro</td>
	<td class="campor" width="34%" align="center">caixa.gov.br</td>
	<td class="campor" width="33%"></td>
</tr>
</table><%
rs.close
end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>