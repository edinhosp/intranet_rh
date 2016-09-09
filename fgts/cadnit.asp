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
<title>Cadastro NIT - Documento de Cadastramento do Trabalhador no PIS</title>
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
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

espacamento=5
if request.form="" then
sql="select p.chapa, p.nome from corporerm.dbo.pfunc p where p.chapa<'10000' and p.codtipo='N' order by p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="form" action="cadnit.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleção de Funcionário para emissão do cadastro NIT</td>
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
<a href="dct.asp">Formulário Antigo DCT</a>
</form>

<%
else
sql="select chapa, f.nome, pispasep, admissao, f.dtnascimento, f.sexo, f.carteiratrab, f.seriecarttrab, f.ufcarttrab, f.rua, f.numero, f.complemento, " & _
"f.bairro, f.cidade, f.estado, f.cep, f.codsecao, mae, f.tituloeleitor, f.naturalidade, f.estadonatal, f.cpf, f.cartidentidade, f.orgemissorident, f.ufcartident, " & _
"f.zonatiteleitor, f.secaotiteleitor, f.nacionalidade, pai, f.grauinstrucao, instrucao, estcivil, " & _
"p.datachegada, p.corraca, f.dtemissaoident, f.dtcarttrab, f.telefone1, f.telefone2, f.email " & _
"from qry_funcionarios f inner join corporerm.dbo.ppessoa p on p.codigo=f.codpessoa " & _
"where chapa='" & request.form("chapa") & "' "
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
	admissao=rs("admissao")                :    pai=rs("pai")
	instrucao=rs("instrucao") : estadocivil=rs("estcivil") : datachegada=rs("datachegada") : cor=rs("corraca")
	dtemissaoident=rs("dtemissaoident") : dtcarttrab=rs("dtcarttrab")
	telefone1=rs("telefone1") : telefone2=rs("telefone2") : email=rs("email")
else 
	campus="03"
	nome="" : dtnascimento="   /    /     " : sexo="" : mae="" : naturalidade="" : estadonatal=""
	nacionalidade="" : carteiratrab="" : seriecarttrab="" : ufcarttrab="" : cpf="" : cartidentidade=""
	orgemissorident="" : ufcartident="" : tituloeleitor="" : zonatiteleitor="" : secaotiteleitor=""
	rua="" : numero="" : complemento="" : bairro="" : cidade="" : estado="" : cep=""
	admissao="" : pai="" : instrucao="" : estadocivil="" : datachegada="   /    /     " : cor="9"
	dtemissaoident="   /    /     " : dtcarttrab="   /    /     "
	telefone1="" : telefone2="" : email=""
end if
if campus="01" then cnpj="73.063.166/0001-20":endereco="Rua Narciso Sturlini, 883<br>Jd. Umuarama - CEP 06018-903<br>OSASCO - SP"
if campus="03" then cnpj="73.063.166/0003-92":endereco="Av. Franz Voegelli, 300<br>Vila Yara - CEP 06020-190<br>OSASCO - SP"
if campus="04" then cnpj="73.063.166/0004-73":endereco="Av. Franz Voegelli, 1743<br>Jd. Wilson - CEP 06020-190<br>OSASCO - SP"
if sexo="M" then sexoM="X" else sexoM=""
if sexo="F" then sexoF="X" else sexoF=""
largura=690
%>
<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo width=105 rowspan=2 valign=top><img src="../images/rdt_caixa.gif" height="33" border="0"></td>
	<td class="campop" valign=middle colspan=2 ><b>&nbsp;&nbsp;&nbsp;Cadastro NIS - Documento de Cadastramento</td>
</tr>
<tr>
	<td class=campo width=<%=largura-80-105%></td>
	<td class=campo width=80 style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid">Grau de sigilo<br>#00</td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<!-- 2 identificação -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;01 [ x ] CNPJ [&nbsp;&nbsp;&nbsp;] CEI</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;02 Nome do Empregador</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;03 Data de Vínculo</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%=cnpj%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=admissao%>"></td>
</tr>
<tr><td class="campor" align="center" colspan=3>Os campos 01, 02 e 03 são de preenchimento exclusivo para cadastramento do Trabalhador</td></tr>
<tr><td class="campor" height=5></td></tr>
</table>

<!-- 3 dados cadastrais -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;04 Nome</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="<%=nome%>"></td>
</tr>
<tr><td class="campor" colspan=1 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;04 Nome - continuação</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;05 Data de nascimento</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;06 Sexo</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="60" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%=dtnascimento%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;[<input class=form_input type="text" name="r" style="text-align:center" size="1" value="<%=sexoM%>">] M
		&nbsp;[<input class=form_input type="text" name="r" style="text-align:center" size="1" value="<%=sexoF%>">] F
		</td>
</tr>
<tr><td class="campor" colspan=3 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;07 Nome do Pai</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="<%=pai%>"></td>
</tr>
<tr><td class="campor" colspan=1 height=5></td></tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;08 Nome do Mãe</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="<%=mae%>"></td>
</tr>
<tr><td class="campor" colspan=1 height=5></td></tr>
</table>

<%
if nacionalidade="10" then nacB="X" else nacB="&nbsp;"
if nacionalidade="20" then nacBN="X" else nacBN="&nbsp;"
if nacionalidade>"20" then nacE="X" else nacE="&nbsp;"
if nacionalidade="10" then opais="Brasil" else opais=""
%>
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;09 Nacionalidade</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;10 País de Origem</td>
	<td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;11 UF e Município de Nascimento</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		[<input class=form_input type="text" name="r" style="text-align:center" size="1" value="<%=nacB%>">] Brasileira &nbsp; 
		[<input class=form_input type="text" name="r" style="text-align:center" size="1" value="<%%>">] Brasileiro nascido no exterior<br>
		[<input class=form_input type="text" name="r" style="text-align:center" size="1" value="<%=nacBN%>">] Naturalizado &nbsp; 
		[<input class=form_input type="text" name="r" style="text-align:center" size="1" value="<%=nacE%>">] Estrangeira
	
		</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=opais%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="2" value="<%=estadonatal%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%=naturalidade%>"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>
<%
if len(tituloeleitor)>0 then
	titulo=left(tituloeleitor,len(tituloeleitor)-2)
	dv=right(tituloeleitor,2)
else
	titulo="":dv=""
end if

sql2="select descricao from corporerm.dbo.pcorraca where codcliente=" & cor
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
cor=rs2("descricao"):rs2.close
%>
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;12 Cor</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;13 Estado Civil</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;14 Nível de Instrução</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;15 Data de Chegada no Brasil</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=cor%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%=estadocivil%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="30" value="<%=instrucao%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=datachegada%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" rowspan=2 valign=top><b>&nbsp;16 CPF</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=5><b>&nbsp;17 Identidate</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;17.1 Número</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;17.2 Complemento</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;17.3 UF</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;17.4 Emissor</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;17.5 Data Emissão</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=cpf%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=cartidentidade%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="2" value="<%=ufcartident%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=orgemissorident%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=dtemissaoident%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=4><b>&nbsp;18 CTPS</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;18.1 Número</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;18.2 Série</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;18.3 UF</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;18.4 Data emissão</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=carteiratrab%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=seriecarttrab%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="2" value="<%=ufcarttrab%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="60" value="<%=dtcarttrab%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=3><b>&nbsp;19 Certidão Civil</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;19.1 Tipo</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;19.2 Data de Emissão</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;19.3 Termo/Matrícula</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=5><b>&nbsp;19 Certidão Civil - continuação</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;19.4 Livro</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;19.5 Folha</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;19.6 Cartório</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;19.7 UF</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;19.8 Município</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="25" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=6><b>&nbsp;20 Passaporte</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;20.1 Número</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;20.2 Emissor</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;20.3 UF</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;20.4 Data emissão</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;20.5 Data validade</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;20.6 País de emissão</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=3><b>&nbsp;21 Título de Eleitor</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=2><b>&nbsp;22 Portaria Naturalização</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;21.1 Número</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;21.2 Zona</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;21.3 Seção</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;22.1 Número</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;22.2 Data naturalização</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=tituloeleitor%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=zonatiteleitor%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=secaotiteleitor%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=5><b>&nbsp;23 RIC</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;23.1 Número</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;23.2 UF</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;23.3 Orgão Emissor</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;23.4 Município</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;23.5 Data Expedição</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=4><b>&nbsp;24 Endereço</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;24.1 Tipo</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;24.2 CEP</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;24.3 UF</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;24.4 Município</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;[<input class=form_input type="text" name="r" style="text-align:center" size="1" value="<%%>">] Comercial
		&nbsp;[<input class=form_input type="text" name="r" style="text-align:center" size="1" value="X">] Residencial
		</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=cep%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=estado%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="40" value="<%=cidade%>"></td>
</tr>
<tr><td class="campor" colspan=4 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;24.5 Bairro</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;24.6 Logradouro</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;24.7 Nº</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;24.8 Complemento</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%=bairro%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="30" value="<%=rua%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=numero%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=complemento%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=2><b>&nbsp;25 Caixa Postal</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000" colspan=4><b>&nbsp;26 Telefone</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;25.1 Número</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;25.2 CEP</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;26.1 DDD</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;26.2 Fixo</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;26.3 DDD</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;26.4 Celular</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%="011"%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=telefone1%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%="011"%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		&nbsp;<input class=form_input type="text" name="razao" size="10" value="<%=telefone2%>"></td>
</tr>
<tr><td class="campor" colspan=5 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000"><b>&nbsp;27 Email</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="<%=email%>"></td>
</tr>
<tr><td class="campor" colspan=1 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-left:0px solid #000000;border-right:0px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" style="font-size:12px" name="razao" size="80" value="<%="OSASCO, " & day(now) & " DE " & ucase(monthname(month(now))) & " DE " & year(now)%>"></td>
</tr>
<tr>
	<td class=campo style="border-left:0px solid #000000;border-right:0px solid #000000">&nbsp;Local/Data</td>
</tr>
<tr><td class="campor" colspan=1 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-bottom:1px solid #000000" width="45%" height="25"></td>
	<td class=campo style="border-bottom:0px solid #000000" width="10%"></td>
	<td class=campo style="border-bottom:1px solid #000000" width="45%"></td>
</tr>
<tr>
	<td class=campo style="" width="45%" align="center">Assinatura do solicitante</td>
	<td class=campo style="" width="10%"></td>
	<td class=campo style="" width="45%" align="center">Assinatura Empregado CAIXA - Sob carimbo</td>
</tr>
<tr>
	<td class=campo colspan=3 align="left">31.445 v002 micro</td>
</tr>
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