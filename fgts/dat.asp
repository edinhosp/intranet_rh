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
<title>DAT - Documento de Atualização de Dados do Trabalhador</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"><!--
function nome1() {	form.chapa.value=form.nome.value; }
function chapa1() {	form.nome.value=form.chapa.value; }
--></script>
</head>
<body style="margin-left:20px">
<%
dim conexao, rs, npis(11), dnasc(8)
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
<form name="form" action="dat.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleção de Funcionário para emissão de DAT</td>
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
</form>

<%
else
sql="select chapa, nome, pispasep, admissao, dtnascimento, sexo, carteiratrab, seriecarttrab, ufcarttrab, rua, numero, complemento, " & _
"bairro, cidade, estado, cep, codsecao, secao, mae, tituloeleitor, naturalidade, estadonatal, cpf, cartidentidade, orgemissorident, ufcartident, " & _
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
	pispasep=rs("pispasep")                :        secao=rs("secao")
	dtnascimento=numzero(day(dtnascimento),2) & numzero(month(dtnascimento),2) & numzero(year(dtnascimento),4)
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
	pispaep=""          :   secao=""
end if
if campus="01" then cnpj="73.063.166/0001-20":endereco="Rua Narciso Sturlini, 883<br>Jd. Umuarama - CEP 06018-903<br>OSASCO - SP"
if campus="03" then cnpj="73.063.166/0003-92":endereco="Av. Franz Voegelli, 300<br>Vila Yara - CEP 06020-190<br>OSASCO - SP"
if campus="04" then cnpj="73.063.166/0004-73":endereco="Av. Franz Voegelli, 1743<br>Jd. Wilson - CEP 06020-190<br>OSASCO - SP"
largura=690
for a=1 to 11
	npis(a)=mid(pispasep,a,1)
next

for a=1 to 8
	dnasc(a)=mid(dtnascimento,a,1)
next
%>
<br><br>
<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class=campo width=130 valign=top><img src="../images/rdt_caixa.gif" height="33" border="0"></td>
	<td class="campop" width=520 valign=top align="center" ><b>DAT - Documento de Atualização de Dados do Trabalhador</td>
	<td class="campop" width=40 valign=top align="right">PIS</td>
</tr>
</table>


<!-- -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class=campo valign=top width=80%>
<!-- -->

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura*.68%>" style="border-collapse: collapse">
<tr>
	<td class="campop" width=230 rowspan=2 style="" valign="middle">&nbsp;Preenchimento obrigatório <img src="../images/setanext1.gif" width="12" height="12" border="0" alt=""></td>
	<td class=fundor colspan=11 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;01 - Inscrição</td>
</tr>
<tr>
	<%for a=1 to 11%>
	<td class=fundop align="center" width=20 style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000"><b><%=npis(a)%></td>
	<%next%>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura*.68%>" style="border-collapse: collapse">
<tr><td class=campo><b>Preencher conforme instrução no verso</td></tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;03 - Nome do trabalhador</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value="<%=nome%>"></td>
</tr>
<tr>
	<td class=campo height=25 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="80" value=""></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura*.68%>" style="border-collapse: collapse">
<tr>
	<td class="campor" colspan=8 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;04 - Data de nascimento</td>
	<td class="campor" width=5>&nbsp;</td>
	<td class="campor" style="">&nbsp;05 - Sexo</td>
</tr>
<tr>
	<%for a=1 to 8%>
	<td class=campo align="center" width=20 style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000"><b>
	<input class=form_input type="text" name="dtnasc<%=a%>" size="1" value="<%=dnasc(a)%>"></td>
	<%next%>
	<td class="campor" width=5>&nbsp;</td>

	<%
	t1="¨":t2="¨"
	if sexo="F" then t1="þ"
	if sexo="M" then t2="þ"
	t1="¨":t2="¨"
	%>
	<td class=campo style="">&nbsp;
	<font style="font-family:Wingdings;font-size:16pt"><%=t2%><font style="font-family:Tahoma;font-size:8pt">Masculino&nbsp;
	<font style="font-family:Wingdings;font-size:16pt"><%=t1%><font style="font-family:Tahoma;font-size:8pt">Feminino
	</td>
</tr>
<tr><td class="campor" colspan=10 height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura*.68%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;06 - Nome da mãe</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="mae" size="80" value="<%=mae%>"></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<!-- -->
</td><td class=campo valign=top width=20%>
<!-- -->
<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" width="<%=largura*.3%>" style="border-collapse: collapse">
<tr>
	<td class=campo width=100% height=188 valign=top style="border-left: 1px solid;border-bottom: 1px solid;border-right: 1px solid">02 - Carimbo da Agência Receptora<br>&nbsp;&nbsp;&nbsp;(Norma CSA/CIEF nº 047)
	<br>&nbsp;<br><br><br><br><br></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<!-- -->
</td></tr></table>
<!-- -->


<!-- 2 identificação -->
<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;07 - Município de nascimento</td>
	<td class="campor" rowspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
	<td class="campor" rowspan=2 width=5>&nbsp;</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;08 - Código</td>
	<td class="campor" rowspan=2 width=5>&nbsp;</td>
	<td class="campor" colspan=3 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;09 - Carteira de trabalho</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Nacionalidade</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Número</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Série</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;UF</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="40" value="<%=naturalidade%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="2" value="<%=estadonatal%>"></td>
	<td class="campor" width=5>&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="2" value="<%=nacionalidade%>"></td>
	<td class="campor" width=5>&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=carteiratrab%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="2" value="<%=seriecarttrab%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="2" value="<%=ufcarttrab%>"></td>
	</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" colspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;10 - CPF</td>
	<td class="campor" rowspan=2 width=5>&nbsp;</td>
	<td class="campor" colspan=2 valign=top style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;11 - Carteira de identidade</td>
	<td class="campor" rowspan=2 width=5>&nbsp;</td>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;12 - Título de eleitor</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Número</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;DV</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Número</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Emissor</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;Número</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;DV</td>
</tr>
<%
if isnull(cpf) or cpf="" then cpf="  " else cpf=cpf
cpf1=left(cpf,len(cpf)-2)
cpf2=right(cpf,2)
if isnull(tituloeleitor) or tituloeleitor="" then tituloeleitor="  " else tituloeleitor=tituloeleitor
tit1=left(tituloeleitor,len(tituloeleitor)-2)
tit2=right(tituloeleitor,2)
%>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=cpf1%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=cpf2%>"></td>
	<td class="campor" width=5>&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=cartidentidade%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="15" value="<%=orgemissorident%>"></td>
	<td class="campor" width=5>&nbsp;</td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="20" value="<%=tit1%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="razao" size="5" value="<%=tit2%>"></td>
	</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;13 - Endereço do trabalhador (logradouro, número e complemento</td>
</tr>
<tr><%endereco=rua & " " & numero & " " & complemento%>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="mae" size="80" value="<%=endereco%>"></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;14 - Bairro</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;15 - Município</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;16 - UF</td>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;17 - CEP</td>
</tr>
<tr>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="mae" size="30" value="<%=bairro%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="mae" size="30" value="<%=cidade%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="mae" size="2" value="<%=estado%>"></td>
	<td class=campo style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	&nbsp;<input class=form_input type="text" name="mae" size="7" value="<%=cep%>"></td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr>
	<td class="campor" width=150 height=40>&nbsp;</td>
	<td class=campo style="border-bottom:1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class="campor" width=5>&nbsp;</td>
	<td class="campor" style="">&nbsp;18 - Assinatura do trabalhador</td>
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<DIV style="page-break-after:always"></DIV>

<table border="0" cellpadding="2" cellspacing="0" width="<%=largura%>" style="border-collapse: collapse">
<tr><td class=campo colspan=2><b>INSTRUÇÕES DE PREENCHIMENTO</td></tr>
<tr>
	<td class=campo colspan=2><b>Preencher obrigatoriamente, o campo 1 e os os campos a serem ATUALIZADOS, conforme instruções abaixo:</td></tr>
<tr>
	<td class=campo width=20% valign=top><b>CAMPO 01</td><td class=campo width=80%>Número de inscrição do trabalhador no PIS.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 03</td><td class=campo>Nome do empregado, trabalhador avulso ou temporário, sem abreviações, se possível.<br>Não abreviar o primeiro e o último nome.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 04</td><td class=campo>Data de nascimento do trabalhador.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 05</td><td class=campo>assinale se Masculino ou Feminino.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 06</td><td class=campo>Nome da mãe do trabalhador, sem abreviações, se possível.<br>Se desconhecida, preencher com a expressão "IGNORADA".<br>Não abreviar o primeiro e o último nome.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 07</td><td class=campo>Nome do munícipio de nascimento do trabalhador, incluindo a sigla da Unidade da Federação. Somente para brasileiro.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 08</td><td class=campo>Nacionalidade do trabalhador. Preenchida conforme tabela:
		<table border="0" cellpadding="0" cellspacing="0" width="500" style="border-collapse: collapse">
		<tr><td class="campor">10 - BRASILEIRA</td>      <td class="campor">20 - NATURALIZADO</td>  <td class="campor">21 - ARGENTINA</td></tr>
		<tr><td class="campor">22 - BOLIVIANA</td>       <td class="campor">23 - CHILENA</td>       <td class="campor">24 - PARAGUAIA</td></tr>
		<tr><td class="campor">25 - URUGUAIA</td>        <td class="campor">30 - ALEMÃ</td>         <td class="campor">31 - BELGA</td></tr>
		<tr><td class="campor">32 - BRITÂNICA</td>       <td class="campor">34 - CANADENSE</td>     <td class="campor">35 - ESPANHOLA</td></tr>
		<tr><td class="campor">36 - NORTE AMERICANA</td> <td class="campor">37 - FRANCESA</td>      <td class="campor">38 - SUIÇA</td></tr>
		<tr><td class="campor">39 - ITALIANA</td>        <td class="campor">41 - JAPONESA</td>      <td class="campor">42 - CHINESA</td>
		<tr><td class="campor">43 - COREANA</td>         <td class="campor">45 - PORTUGUESA</td>    <td class="campor">48 - OUTRAS LATINO AMERICANAS</td></tr>
		<tr><td class="campor">49 - OUTRAS ASIÁTICAS</td><td class="campor">50 - OUTRAS</td>        <td class="campor">&nbsp;</td></tr>
		</table>
	</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 09</td><td class=campo>Carteira de Trabalho e Previdência Social do trabalhador, com número, série e sigla da Unidade da Federação emissora.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 10</td><td class=campo>Número e dígito verificador do CPF do trabalhador.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 11</td><td class=campo>Número da Carteira de Identidade do trabalhador e sigla do orgão emissor:
		<table border="0" cellpadding="0" cellspacing="0" width="250" style="border-collapse: collapse">
			<tr><td class="campor">ORGÃO EMISSOR</td>                   <td class="campor">PREENCHIMENTO</td></tr>
			<tr><td class="campor">Ministério da Marinha</td>           <td class="campor">MR</td></tr>
			<tr><td class="campor">Ministério da Aeronáutica</td>       <td class="campor">AE</td></tr>
			<tr><td class="campor">Ministério do Exército</td>          <td class="campor">EX</td></tr>
			<tr><td class="campor">Carteira modelo 19 (estrangeiro)</td><td class="campor">DE</td></tr>
			<tr><td class="campor">Secretaria de Segurança Pública</td> <td class="campor">Sigla da UF</td></tr>
			<tr><td class="campor">Outros emissores</td>                <td class="campor">OE</td></tr>
		</table>
	</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 12</td><td class=campo>Número e digito verificador do Título de Eleitor do trabalhador.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 13 a 17</td><td class=campo>Endereço do trabalhador, contendo logradouro, número e complemento (apartamento, bloco, quadra etc), bairro, município, UF e CEP.</td></tr>
<tr>
	<td class=campo valign=top><b>CAMPO 18</td><td class=campo>Assinatura do trabalhador</td></tr>
<tr>	<td class="campor" colspan=2 valign=top align="right"><%=lcase(nome & " - " & secao)%></td></tr>
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