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
<title>DCT - Documento de Cadastramento do Trabalhador no PIS</title>
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
<form name="form" action="dct.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleção de Funcionário para emissão de DCT</td>
</tr>
<tr>
	<td class=campo>Funcionário</td>
	<td class=campo><input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()"></td>
	<td class=campo>
		<select name="nome" class=a onchange="nome1()">
		<option value="0"> Selecione o funcionário</option>
		<option value="00000">Formulário em Branco</option>
<%rs.movefirst:do while not rs.eof%>
		<option value="<%=rs("chapa")%>"> <%=rs("nome")%></option>
<%rs.movenext:loop
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
<a href="cadnit.asp">Formulário Novo: Cadastro NIT (a partir de Maio/2012)</a>
</form>

<%
else
sql="select chapa, nome, pispasep, admissao, dtnascimento, sexo, carteiratrab, seriecarttrab, ufcarttrab, rua, numero, complemento, " & _
"bairro, cidade, estado, cep, codsecao, mae, tituloeleitor, naturalidade, estadonatal, cpf, cartidentidade, orgemissorident, ufcartident, " & _
"tituloeleitor, zonatiteleitor, secaotiteleitor, nacionalidade " & _
"from qry_funcionarios where chapa='" & request.form("chapa") & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then 
	campus=left(rs("codsecao"),2) 
	nome=rs("nome")                        :	dtnascimento=rs("dtnascimento")
	sexo=rs("sexo")                        :	mae=rs("mae")
	naturalidade=rs("naturalidade")        :	estadonatal=rs("estadonatal")
	nacionalidade=rs("nacionalidade")      :	carteiratrab=rs("carteiratrab")
	seriecarttrab=rs("seriecarttrab")      :	ufcarttrab=rs("ufcarttrab")
	cpf=rs("cpf")                          :	cartidentidade=rs("cartidentidade")
	orgemissorident=rs("orgemissorident")  :	ufcartident=rs("ufcartident")
	tituloeleitor=rs("tituloeleitor")      :	zonatiteleitor=rs("zonatiteleitor")
	secaotiteleitor=rs("secaotiteleitor")  :	rua=rs("rua")
	numero=rs("numero")                    :	complemento=rs("complemento")
	bairro=rs("bairro")                    :	cidade=rs("cidade")
	estado=rs("estado")                    :	cep=rs("cep")
else 
	campus="03"
	nome=""             :	dtnascimento=""
	sexo=""             :	mae=""
	naturalidade=""     :	estadonatal=""
	nacionalidade=""    :	carteiratrab=""
	seriecarttrab=""    :	ufcarttrab=""
	cpf=""              :	cartidentidade=""
	orgemissorident=""  :	ufcartident=""
	tituloeleitor=""    :	zonatiteleitor=""
	secaotiteleitor=""  :	rua=""
	numero=""           :	complemento=""
	bairro=""           :	cidade=""
	estado=""           :	cep=""
end if
if campus="01" then cnpj="73.063.166/0001-20":endereco="Rua Narciso Sturlini, 883<br>Jd. Umuarama - CEP 06018-903<br>OSASCO - SP"
if campus="03" then cnpj="73.063.166/0003-92":endereco="Av. Franz Voegelli, 300<br>Vila Yara - CEP 06020-190<br>OSASCO - SP"
if campus="04" then cnpj="73.063.166/0004-73":endereco="Av. Franz Voegelli, 1743<br>Jd. Wilson - CEP 06020-190<br>OSASCO - SP"
largura=690
%>
<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" width="<%=largura%>" height=150 style="border-collapse: collapse">
<tr>
	<td class=campo width=105 valign=top><br><br><img src="../images/rdt_caixa.gif" height="33" border="0"></td>
	<td class="campop" width=140 valign=top ><b>DCT - Documento de Cadastramento do Trabalhador no PIS</td>
	<td class="campor" width=240 valign=top style="border-left: 1px solid;border-bottom: 1px solid"> 01 - Carimbo padronizado do CGC ou <br>matrícula no Cadastro Específico do INSS-CEI
		<div align="center"><center>
		<table border="0" cellpadding="0" width="240" cellspacing="0">
		<tr><td width="1" style="border-left-style: solid; border-left-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240" rowspan="2">
				<p align="center"><b><font size="4" color="#808080"><%=cnpj%></font></b></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
		<tr><td width="1"></td><td width="1"></td></tr>
		<tr><td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td><td width="1"></td></tr>
		<tr><td width="1"></td><td width="240" align="center">
				<b><font color="#808080">FUNDAÇÃO INSTITUTO DE<br>ENSINO PARA OSASCO</font></b></td>
			<td width="1"></td></tr>
		<tr><td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td><td width="1"></td></tr>
		<tr><td width="1">&nbsp;</td><td width="240" rowspan="2" align="center">
				<font color="#808080"><%=endereco%></font></td><td width="1"></td></tr>
		<tr><td width="1" style="border-left-style: solid; border-left-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
		</table></center></div>
	</td>
	<td class=fundor width=205 valign=top style="border-left: 1px solid;border-bottom: 1px solid;border-right: 1px solid"> Para uso exclusivo da CAIXA<br>Carimbo da Agência Receptora<br>Norma CSA/CIEF nº 047</td>
</tr>
</table>

<!-- 2 identificação -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class="campor"><b>2 - Identificação do Empregador/Sindicato</td></tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CGC/CEI</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nome</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%=cnpj%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO"></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Endereço</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="100" value="<%=replace(endereco,"<br>"," - ")%>"></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Telefone</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;FAX</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="25" value="3651-9905"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="25" value="3651-9987"></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<!-- 3 dados cadastrais -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class="campor" colspan=1><b>IDENTIFICAÇÃO DO TRABALHADOR</td></tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;03 - Nome do Trabalhador</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="<%=nome%>"></td>
</tr>
<tr><td class="campor" colspan=1 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;04 - Data de nascimento</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;05 - Sexo</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;06 - Nome da mãe</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%=dtnascimento%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=sexo%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="60" value="<%=mae%>"></td>
</tr>
<tr><td class="campor" colspan=3 height=5></td></tr>
</table>

<!-- -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class=campo valign=top width=80%>
<!-- -->

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura*.76%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;07 - Município de nascimento</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;08 - Cód. Nasc.</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="50" value="<%=naturalidade%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=estadonatal%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=nacionalidade%>"></td>
</tr>
<tr><td class="campor" colspan=3 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura*.76%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;09 - Carteira de trabalho nº</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Série</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;10 - CPF</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=carteiratrab%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=seriecarttrab%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=ufcarttrab%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%=cpf%>"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>

<!-- -->
</td><td class=campo valign=top width=20%>
<!-- -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura*.23%>" style="border-collapse: collapse">
<tr>
	<td class=fundor style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>Para uso exclusivo da CAIXA</td>
</tr>
<tr>
	<td class=fundo style="border-left:1px solid #000000;border-right:1px solid #000000;">
	<font style="font-family:Webdings;font-size:12pt">c <font style="font-family:Tahoma;font-size:8pt">Solicitação atendida</font>
	<br><font style="font-family:Webdings;font-size:12pt">c <font style="font-family:Tahoma;font-size:8pt">Preencimento incorreto</font>
	</td>
</tr>
<tr><td class=fundor style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000" colspan=1 height=15></td></tr>
</table>
<!-- -->
</td></tr></table>
<!-- -->

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;11 - Carteira de identidade nº</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Emissor</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;12 - Título de eleitor nº</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;DV</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Inscrição</td>
</tr>
<%
if len(tituloeleitor)>0 then
	titulo=left(tituloeleitor,len(tituloeleitor)-2)
	dv=right(tituloeleitor,2)
else
	titulo="":dv=""
end if
%>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%=cartidentidade%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=orgemissorident & " " & ufcartident%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%=titulo%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=dv%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=zonatiteleitor & "/" & secaotiteleitor%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;13 - Endereço do trabalhador</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="<%=rua & " " & numero & " " & complemento%>"></td>
</tr>
<tr><td class="campor" colspan=1 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Bairro</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Município</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;CEP</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="30" value="<%=bairro%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%=cidade%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=estado%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=cep%>"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
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