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
<title>Solicitação de Extrato de Conta Vinculada do FGTS</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"><!--
function nome1() {	form.chapa1.value=form.nome1.value; }
function chapa1() {	form.nome1.value=form.chapa1.value; }
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
if request.form("B1")="" then
'sql="select p.chapa, p.nome from corporerm.dbo.pfunc p where p.chapa<'10000' and p.codtipo='N' and codsituacao<>'D' order by p.nome "
'rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="form" action="solicitacaoextrato.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=2>Seleção de Funcionário para emissão de Solicitação de Extrato</td>
</tr>
<%for a=1 to 5%>
<tr>
	<td class=campo><input type='text' name='chapa<%=a%>' size='6' maxlength='5' value='<%=request.form("chapa"&a)%>' onchange='javascript:submit()'>
	</td>
	<td class="campol" width=300>
<%
if request.form("chapa"&a)<>"" then
	sql="select nome from corporerm.dbo.pfunc where chapa='" & request.form("chapa"&a) & "' "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then response.write " <font size=2><b>" & rs("nome") & "</b>"
	rs.close
end if
%>	
	</td>
</tr>
<%
next
%>

<tr>
	<td class=campo colspan=3>&nbsp;
		<input type="submit" value="Visualizar" class=button name="B1">
	</td>
</tr>
</table>
</form>

<%
else
largura=690 '650
larg2=685 '445
%>
<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo valign=top width=175><img src="../images/rdt_caixa.gif" height="33" border="0"></td>
	<td class="campop" valign=middle align="left"><b>Solicitação de Extrato de Conta Vinculada do FGTS</td>
	<td width=70>&nbsp;</td>
</tr>
<tr>
	<td colspan=2 >&nbsp;</td>
	<td class=campo height=35 width=70 style="border-left: 1px solid;border-bottom: 1px solid;border-right: 1px solid" valign=top>Grau de sigilo<br>&nbsp;</td>
</tr>
<tr><td class="campor" height=5 colspan=3></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campop">
	À<br>CAIXA ECONOMICA FEDERAL<br><input type="text" class=form_input10 size=40 value="Ag. Bela Vista / SP">	
	<br><br>Senhor(a) Gerente
	</td>
</tr>
<tr><td class="campor" height=5 colspan=1></td></tr>
</table>

<!-- 1 solicitação -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=larg2%>" style="border-collapse: collapse">
<tr><td class="campop" height=20 colspan=8><b>1 - Solicito o fornecimento de extrato da(s) conta(s) FGTS abaixo relacionadas:</td></tr>

<tr><td class="campop" height=22 width=10 style="border:1px solid #000000;"><input class=form_input style="text-align:center" type="text" name="a1" size="1" value=""></td>
	<td class="campop" style="border-left:0px solid #000000;border-right:0px solid #000000">&nbsp;Empresa</td>
	<td class="campop" width=15 style="border:1px solid #000000;"><input class=form_input style="text-align:center" type="text" name="a1" size="1" value=""></td>
	<td class="campop" style="border-left:0px solid #000000;border-right:0px solid #000000">&nbsp;Fins Rescisórios</td>
	<td class="campop" width=15 style="border:1px solid #000000;"><input class=form_input style="text-align:center" type="text" name="a1" size="1" value=""></td>
	<td class="campop" style="border-left:0px solid #000000;border-right:0px solid #000000">&nbsp;Analítico</td>
	<td class="campop" width=15 style="border:1px solid #000000;"><input class=form_input style="text-align:center" type="text" name="a1" size="1" value=""></td>
	<td class="campop" style="border-left:0px solid #000000;border-right:0px solid #000000">&nbsp;Simples Conferência</td>
</tr>
<tr><td class="campor" colspan=8 height=10></td></tr>
</table>

<!-- 1.1 dados do empregador -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=larg2%>" style="border-collapse: collapse">
<tr><td class="campop" height=20 colspan=2><b>1.1 - Dados do Empregador</td></tr>

<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Razão Social do Empregador:</td>
	<td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ/CEI:</td>
</tr>

<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input10 type="text" name="razao" size="50" value="FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO"></td>
	<td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input10 type="text" name="cnpj" size="25" value="73.063.166/0001-20"></td>
	</tr>
<tr><td class="campor" colspan=2 height=10></td></tr>
</table>

<!-- 1.2 dados do trabalhador -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=larg2%>" style="border-collapse: collapse">
<tr><td class="campop" colspan=3><b>1.2 - Dados do Trabalhador</b> <span style="font-size:8pt">(Para solicitação pelo empregador dispensa o preenchimento do campo CNPJ)</span></td></tr>
<%
dim chapa(5)
for a=1 to 5
if request.form("chapa"&a)="" then chapa(a)="00000" else chapa(a)=request.form("chapa"&a)
sql="select f.chapa, f.nome, f.dataadmissao, f.pispasep, f.codsecao " & _
"from corporerm.dbo.pfunc f where f.chapa='" & request.form("chapa"&a) & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	nome=rs("nome")
	campus=left(rs("codsecao"),2)
	if campus="01" then cnpj="73.063.166/0001-20"
	if campus="03" then cnpj="73.063.166/0003-92" 
	if campus="04" then cnpj="73.063.166/0004-73" 
	pispasep=rs("pispasep")
	dataadmissao=rs("dataadmissao")
else
	nome=""
	cnpj=""
	pispasep=""
	dataadmissao=""
end if
%>
<tr><td class=campo colspan=3 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome</td>
</tr>
<tr><td class=campo colspan=3 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input10 type="text" name="nome<%=a%>" size="70" value="<%=nome%>"></td>
</tr>

<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CNPJ Empregador</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Inscrição PIS/PASEP</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Data Admissão</td>
</tr>
<tr><td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input10 type="text" name="razao" size="40" value="<%=cnpj%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input10 type="text" name="razao" size="20" value="<%=pispasep%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input10 type="text" name="razao" size="20" value="<%=dataadmissao%>"></td>
</tr>
<tr><td class="campor" colspan=3 height=10></td></tr>
<%
rs.close
next
%>
<tr><td class="campor" colspan=3 height=1></td></tr>
</table>

<!-- 2 dados complementares -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=larg2%>" style="border-collapse: collapse">
<tr><td class="campop" colspan=4><b>2 - Seguem dados complementares para contato</td></tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Telefone</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Ramal</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Responsável</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;E-mail</td>
</tr>
<tr>
	<td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input10 type="text" name="telefone" size="15" value="3651-9905"></td>
	<td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input10 type="text" name="ramal---" size="10" value="9957"></td>
	<td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input10 type="text" name="nomeresp" size="15" value="Graziela"></td>
	<td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input10 type="text" name="email---" size="25" value="graziela@unifieo.br"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
<tr>
	<td class=campo colspan=4 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Documentos em anexo:</td>
</tr>
<tr>
	<td class="campop" colspan=4 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input10 type="text" name="cnpj" size="85" value=""></td>
</tr>

<tr><td class="campor" colspan=4 height=12></td></tr>
</table>

<!-- final do formulario -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campop" colspan=2 style="">
		<input class=form_input10 style="border-bottom:1px solid #000000" type="text" name="cidade" size="45" value="<%="OSASCO"%>">
		, &nbsp; <input class=form_input10 style="border-bottom:1px solid #000000;text-align:center" type="text" name="dia" size="5" value="<%=day(now)%>">
		de &nbsp;<input class=form_input10 style="border-bottom:1px solid #000000" type="text" name="mes" size="25" value="<%=ucase(monthname(month(now)))%>">
		de &nbsp;<input class=form_input10 style="border-bottom:1px solid #000000;text-align:center" type="text" name="ano" size="10" value="<%=year(now)%>">
	</td>
</tr>
<tr>
	<td class="campop" style="border-top:0px solid #000000" valign=top>Local/Data</td>
	<td class=campo width=50%>&nbsp;</td>
</tr>
<tr><td class="campor" colspan=2 height=35></td></tr>
<tr>
	<td class="campop" style="border-top:1px solid #000000">Assinatura/CI<br>
	NOME: <input type="text" class=form_input size="50" value="ROGERIO MATEUS DOS SANTOS ARAUJO"><br>
	CPF: <input type="text" class=form_input size="20" value="185.420.058-56">
	</td>
	<td class=campo>&nbsp;</td>
<!--RG 27.831.325-5-->
</tr>

<tr><td class="campor" colspan=3 height=1></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class="campor" colspan=1 height=15></td></tr>
<tr><td colspan=1 align="center">
	<p style="font-size:8pt;margin-top:0px;margin-bottom:0px;"><b>SAC CAIXA:</b> 0800 726 0101 (informações, reclamações, sugestões e elogios)</p>
	<p style="font-size:7pt;margin-top:0px;margin-bottom:0px;"><b>Para pessoas com deficiência auditiva:</b> 0800 726 2492</p>
	<p style="font-size:7pt;margin-top:0px;margin-bottom:0px;"><b>Ouvidoria:</b> 0800 725 7474 (reclamações não solucionadas e denúncias)</p>
	<p style="font-size:7pt;margin-top:0px;margin-bottom:0px;"><b>caixa.gov.br</p>
	</td>
</tr>
<tr><td class="campor" colspan=1 align="left" valign=top>31.446 v002 micro</td></tr>
</table>

<%
end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>