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
conexao.Open application("consql")
sessao=session.sessionid
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

espacamento=5
if request.form="" then
sql="select p.chapa, p.nome from pfunc p where p.chapa<'10000' and p.codtipo='N' order by p.nome "
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
"from pfunc f, ppessoa p where f.codpessoa=p.codigo and f.chapa='" & request.form("chapa") & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
campus=left(rs("codsecao"),2)
if campus="01" then cnpj="73.063.166/0001-20"
if campus="03" then cnpj="73.063.166/0003-92" 
if campus="04" then cnpj="73.063.166/0004-73" 
categoria=numzero(rs("codcategoria"),2)
%>
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top><img src="../images/rdt_caixa.gif" height="33" border="0"></td>
	<td class=campo valign=top><img src="../images/rdt_mtb.jpg"   height="36" border="0"></td>
	<td class="campop" valign=middle align="right" width=400><b>R D T - Retificação de Dados do Trabalhador - FGTS</td>
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

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ/CEI do empregador/contribuinte</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%=cnpj%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="3" value="SP"></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Codigo empregador/contribuinte no FGTS</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Base da conta</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="35" value="06951101389480"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="3" value="SP"></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="445" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Pessoa para Contato</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;DDD/telefone</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="55" value="ROGERIO"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="18" value="(011) 3651-9905"></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="445" style="border-collapse: collapse">
<tr><td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Endereço eletrônico (email p/contato)</td></tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="rogerio@unifieo.br"></td></tr>
<tr><td class="campor" height=5></td></tr>
</table>

</td>
<td valign=top align="right">
	<table border="0" cellpadding="1" cellspacing="0" width="197" style="border-collapse: collapse">
	<tr><td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Protocolo de Recepção</td></tr>
	<tr><td class=campo height=182 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
	</table>
</td>
</tr>
</table>

<!-- 2 identificação -->
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class="campor"><b>2 - IDENTIFICAÇÃO DO TRABALHADOR (preenchimento obrigatório)</td></tr>
<tr><td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome do Trabalhador</td></tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="<%=rs("nome")%>"></td></tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº do PIS/PASEP/inscrição contribuinte individual</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Admissão</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Código do Trabalhador (categorias com FGTS)</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Categoria</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="35" value="<%=rs("pispasep")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="12" value="<%=rs("dataadmissao")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="35" value="<%=rs("contafgts")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=categoria%>"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>

<!-- 3 dados cadastrais -->
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class="campor" colspan=2><b>3 - DADOS CADASTRAIS (preencher somente os campos a serem alterados)</td></tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome do Trabalhador</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;NºPIS/PASEP/inscrição contribuinte individual</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Categoria</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="70" value="<%=rs("nome")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("pispasep")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=categoria%>"></td>
</tr>
<tr><td class="campor" colspan=2 height=5></td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Admissão</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de opção</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de retroação</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de Nascimento</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº CTPS</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Série</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("dataadmissao")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("dtopcaofgts")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=rs("dtnascimento")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="7" value="<%=rs("carteiratrab")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=rs("seriecarttrab")%>"></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="3" value="<%=rs("ufcarttrab")%>"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Movimentação informada</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Movimentação correta</td>
</tr>
<tr>
	<td class="campor" width=100 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data</td>
	<td class="campor" width=50  style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Código</td>
	<td class="campor">&nbsp;</td>
	<td class="campor" width=100 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data</td>
	<td class="campor" width=50  style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Código</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
	<td class="campor">&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>

<!-- 4 dados a retificar -->
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class="campor"><b>4 - Retificação da remuneração sem devolução do FGTS (entre contas do mesmo trabalhador ou entre trabalhadores diversos)</b>
	<br>* Nas guias com recolhimento ao FGTS, as remunerações informadas no campo "PARA" devem ser limitadas aos valores discriminados no campo "DE"</td></tr>
<tr><td class="campor"><b>De:</b> (Preencher com dados informados incorretamente na guia)</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome do trabalhador</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº do PIS/PASEP</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Categoria</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de admissão</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Competência</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:2px solid #000000">&nbsp;Remuneração&nbsp;</td>
</tr>
<%for a=1 to 9%>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=a%>-&nbsp;<input class=form_input type="text" name="razao" size="40" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="11" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="2" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="12" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="8" value="   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:2px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="         ,"></td>
</tr>
<%next%>
<tr><td class="campor" colspan=10 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class="campor"><b>Para:</b> (Preencher com dados corretos para a guia)</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome do trabalhador</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nº do PIS/PASEP</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Categoria</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data de admissão</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Competência</td>
	<td class="campor" valign=top style="border-left:1px solid #000000;border-right:2px solid #000000">&nbsp;Remuneração&nbsp;</td>
</tr>
<%for a=1 to 9%>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=a%>-&nbsp;<input class=form_input type="text" name="razao" size="40" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="11" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="2" value=""></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="12" value="   /   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="8" value="   /"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:2px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="         ,"></td>
</tr>
<%next%>
<tr><td class="campor" colspan=10 height=5></td></tr>
</table>


<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=3><b>Poderão ser exigidos outros documentos, caso a CAIXA julgue necessário.</td></tr>
<tr>
	<td width=48% class="campop" style="">
		&nbsp;<input class=form_input type="text" name="razao" size="50" value="<%="OSASCO, " & day(now) & " DE " & ucase(monthname(month(now))) & " DE " & year(now)%>"></td>
	<td class="campop" style="">&nbsp;</td>
	<td width=48% class="campop" style="">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value=""></td>
</tr>
<tr>
	<td class=campo style="border-top:1px solid #000000">Local e Data</td>
	<td class="campop" style="">&nbsp;</td>
	<td class=campo style="border-top:1px solid #000000">Carimbo e assinatura do responsável<br>
	<input type="text" class=form_input size="60" value="ROGERIO MATEUS DOS SANTOS ARAUJO-CPF 185.420.058-56"></td>
<!--RG 27.831.325-5-->
</tr>
<tr><td class="campor" colspan=3 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=campo colspan=3><b>PARA USO DA CAIXA</td></tr>
<tr><td class=campo colspan=3>Declaro que os documentos apresentados comprovam as alterações solicitadas.</td></tr>
<tr><td class=campo colspan=1></td>
	<td class="campor" colspan=2><!--Novas GFIP/GRFP/GRFC Anexadas--></td></tr>
<tr>
	<td class=campo width=400 style="border-right:0px solid #000000;border-bottom:1px solid #000000">&nbsp;</td>
	<td class=campo width=30 style="border-right:0px solid #000000;border-bottom:0px solid #000000">&nbsp;</td>
	<td class="campor" width=220 style=""><!--0 - Sim<br>1 - Não--></td>
</tr>
<tr><td class="campor" colspan=3>Assinaturo/carimbo do responsável pela conferência</td></tr>
</table>
<%
rs.close
end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>