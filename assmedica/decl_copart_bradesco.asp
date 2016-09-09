<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a85")="N" or session("a85")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Declaração Opcional de Plano de Saúde</title>
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open application("conexao")

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao2

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo")
	if request("codigo")="" then temp=request.form("chapa")
	if isnumeric(temp) then
		info=1
		temp=numzero(temp,5)
		sqlb="AND f.CHAPA='" & temp & "' "
	else
		info=2
		sqlb="AND f.nome like '%" & temp & "%' "
	end if

	sqla="SELECT f.NOME, f.CODSITUACAO, f.CHAPA, f.DATAADMISSAO, f.CODSECAO, f.codsindicato, s.DESCRICAO AS Secao, " & _
	"p.dtnascimento, p.telefone1, p.telefone2, p.telefone3, p.email, p.cpf, p.estadocivil, c.nome as funcao, " & _
	"p.cartidentidade, p.cpf, p.dtnascimento, p.sexo, p.rua, p.numero, p.complemento, p.bairro, p.cidade, p.cep, p.estado, " & _
	"p.telefone1, f.datademissao, f.dtaposentadoria, f.aposentado, f.tipodemissao, p.grauinstrucao " & _
	"FROM corporerm.dbo.PFUNC f, corporerm.dbo.PSECAO s, corporerm.dbo.PPESSOA p, corporerm.dbo.PFUNCAO c " & _
	"WHERE f.CODSECAO=s.CODIGO and p.codigo=f.codpessoa and c.codigo=f.codfuncao "

	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("chapa")=rs("chapa")
	temp=0
	if rs.recordcount>1 then temp=2
else
	temp=1
end if

if temp=1 then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para Declaração opcional de Plano de Saúde - BRADESCO
<form method="POST" action="decl_copart_bradesco.asp">
	<p style="margin-top: 0; margin-bottom: 0">Chapa/Nome <input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>">
	<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>

<%
elseif temp=0 then
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
idade=int((now()-rs("dtnascimento"))/365.25)
if rs("datademissao")="" or isnull(rs("datademissao")) then rsdatademissao=now() else rsdatademissao=rs("datademissao")
%>

<%
sqlplano="SELECT codigo, plano FROM assmed_mudanca " & _
"WHERE chapa='" & rs("chapa") & "' and '" & dtaccess(rsdatademissao) & "' between ivigencia and fvigencia "
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly
plano=rs3("plano")
carteirinha=rs3("codigo")
rs3.close
sqlmae="select nome from corporerm.dbo.pfdepend where chapa='"& rs("chapa") & "' and grauparentesco='7'"
rs3.Open sqlmae, ,adOpenStatic, adLockReadOnly
mae=rs3("nome")
rs3.close

dia1=numzero(day(rs("dtnascimento")),2)
mes1=numzero(month(rs("dtnascimento")),2)
ano1=right(year(rs("dtnascimento")),2)
dtnasc=dia1&mes1&ano1
idade=int((now()-rs("dtnascimento"))/365.25)
dia2=numzero(day(rs("dataadmissao")),2)
mes2=numzero(month(rs("dataadmissao")),2)
ano2=right(year(rs("dataadmissao")),2)
dtadmissao=dia2&mes2&ano2
dia3=numzero(day(rsdatademissao),2)
mes3=numzero(month(rsdatademissao),2)
ano3=right(year(rsdatademissao),2)
dtdemissao=dia3&mes3&ano3
dia4=day(rs("dtaposentadoria")):if dia4="" or isnull(dia4) then dia4="  " else dia4=numzero(dia4,2)
mes4=month(rs("dtaposentadoria")):if mes4="" or isnull(mes4) then mes4="  " else mes4=numzero(mes4,2)
ano4=year(rs("dtaposentadoria")):if ano4="" or isnull(ano4) then ano4="  " else ano4=right(ano4,2)
dtaposent=dia4&mes4&ano4
%>

<%
sqld="SELECT d.chapa, d.dependente, d.nascimento, d.sexo, d.parentesco, d.cpf, d.mae, p.empresa, p.plano " & _
"FROM assmed_dep d, assmed_dep_mudanca p WHERE d.chapa=p.chapa and d.nrodepend=p.nrodepend and p.plano='" & plano & "' " & _
"AND d.chapa='" & rs("chapa") & "' AND '" & dtaccess(rsdatademissao) & "' Between p.ivigencia And p.fvigencia "
rs3.Open sqld, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
%>
<%
rs3.movenext:loop
end if 'rs3.recordcount>0

if rs3.recordcount=0 or rs3.recordcount<4 then
for b=rs3.recordcount+1 to 4
%>
<%
next
end if
rs3.close

'052 desconto co-participação 076 desconto assistencia médica
sqlp="select count(chapa) as vezes from corporerm.dbo.pffinanc where codevento IN ('076','076I','076U','076M') and chapa='" & rs("chapa") & "' "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
meses2=rs3("vezes")
if meses2="" or isnull(meses2) then meses2=0
rs3.close
sqlp="select count(chapa) as vezes from corporerm.dbo.pffinanccompl where codevento IN ('076','076I','076U','076M') and chapa='" & rs("chapa") & "' "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
mes=rs3("vezes")
if mes="" or isnull(mes) then mes=0
meses2=meses2+mes
rs3.close
sqlp="select count(chapa) as vezes from corporerm.dbo.pffinanc where codevento IN ('052','052U') and chapa='" & rs("chapa") & "' "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
meses=rs3("vezes")
if meses="" or isnull(meses) then meses=0
rs3.close
sqlp="select count(chapa) as vezes from corporerm.dbo.pffinanccompl where codevento IN ('052','052U') and chapa='" & rs("chapa") & "' "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
mes=rs3("vezes")
if mes="" or isnull(mes) then mes=0
meses=meses+mes
rs3.close
cano=int((meses+meses2)/12)
cmes=(meses+meses2)-(cano*12)
dini=dtdemissao
sqlp="select max(valor) ultima from corporerm.dbo.pffinanc where codevento in ('052','052U','052I','076','076I','076U','076M') and chapa='" & rs("chapa") & "' " & _
"and /*mescomp=" & month(rs("datademissao")) & " and*/ anocomp=" & year(rs("datademissao")) & " "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
ultimo=rs3("ultima")
if ultimo="" or isnull(ultimo) then ultimo=0
rs3.close
%>


<%
rs.close
set rs=nothing
%>

<div align="center">
<center>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" align="left" width="25%" style="border-bottom:1px solid #000000"><img src="../images/bradesco_saude.jpg" border="0"></td>
	<td class="campop" align="left" valign="top" width="50%" style="border-bottom:1px solid #000000"><b>
	Bradesco Saúde Coletivo Empresarial/SPG<br>
	Formulário de Cancelamento/Alteração de Dados do Segurado e Dependente</b>
	</td>
	<td class="campo" align="left" width="25%" style="border-bottom:1px solid #000000"><img src="../images/bradesco_dental.jpg" border="0"></td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Companhia Seguradora<br>Bradesco Saúde S.A.</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	CNPJ<br>92.693.118/0001-60</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Registro na ANS:<br>005711</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Operadora<br>Odontoprev S.A.</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	CNPJ<br>58.119.199/0001-51</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Registro na ANS:<br>30194-9</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Cia<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Apólice<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Subfatura<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Nome da empresa estipulante / subestipulante<br>&nbsp;		</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="middle" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000">
	Tipo de Movimentação</td>
	<td class="campo" valign="middle" align="left" style="border-bottom:2px solid #000000;border-right:0px solid #000000">
	1 - Cancelamento do titular</td>
	<td class="campo" valign="middle" align="left" style="border-bottom:2px solid #000000;border-right:0px solid #000000">
	2 - Cancelamento de dependente<br>3 - Alteração do titular</td>
	<td class="campo" valign="middle" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000">
	4 - Alteração de dependente<br>5 - 2ª via de cartão</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000">
	Certificado<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:0px solid #000000">
	&nbsp;		</td>
</tr>
<tr>
	<td class="campo" align="left" colspan="6" style="border-bottom:1px solid #000000">
	<b> 1 - Dados do titular (somente preencher caso alteração / cancelamento do titular )</b>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Matrícula Funcional<br>&nbsp;</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Nome<br>&nbsp;		</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="middle" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Preencher somente os<br>dados a serem alterados</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	CNS (Carteira Nacional de Saúde)<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	DNV (Declaração de Nascido Vivo):<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	RIC (Registro de Identificação Civil):<br>&nbsp;		</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Sexo<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	1 - Masculino<br>2 - Feminino</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Data de Nascimento<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Estado Civil<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	1 - Solteiro<br>2 - Casado</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	3 - Viúvo<br>4 - Outros</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	CPF<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	PIS<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Cargo / Ocupação<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Data de Admissão<br>&nbsp;		</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000">
	Nome da Mãe<br>&nbsp;	</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000">
	Nova subfatura<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000">
	Data (Cancelamento ou Alteração)<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000">
	Motivo<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:0px solid #000000">
	Motivo do Cancelamento:<br>	1 - Desistência<br>	2 - Demissão s/justa causa<br>	3 - Aposentadora</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:0px solid #000000">
	&nbsp;<br>	4 - Falecimento<br>	5 - Duplicidade<br>	6 - Portabilidade</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:0px solid #000000">
	&nbsp;<br>	8 - Demissão c/justa causa<br>	9 - Outros</td>
</tr>
<tr>
	<td class="campo" align="left" colspan="7" style="border-bottom:1px solid #000000">
	<b> 2 - Dados do cancelamento (somente preencher em caso de Demissão SEM Justa Causa / Aposentadoria do Titular)</b>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Titular contribuiu <sup><b>1</b></sup> para o pagamento do prêmio?    [&nbsp;&nbsp;] Sim * [&nbsp;&nbsp;] Não </td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Caso positivo, por quanto tempo:       Meses	</td>
</tr>
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" colspan="2">
	Quanto à manutenção no plano oferecido no momento do desligamento ao Titular que contribuiu para o pagamento do prêmio:    [&nbsp;&nbsp;] Aceitou [&nbsp;&nbsp;] Recusou </td>
</tr>
<tr>
	<td class="campo" valign="middle" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" colspan="2" height="50px">
	<b>Assinatura do Segurado Titular:</b> </td>
</tr>
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" colspan="2">
	<sup><b>1</b></sup> Considera-se contribuição o valor pago pelo empregado, a qualquer tempo, para custear parte ou a integralidade do prêmio do seguro, mesmo que no 
	momento da demissão/aposentadoria, o mesmo não esteja contribuindo no pagamento. Não são considerados como contribuições os pagamentos de valores relacionados aos
	dependentes e agregados e à coparticipação ou franquia.</td>
</tr>
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:0px solid #000000" colspan="2">
	* O cancelamento nestes casos está condicionado à entrega da(s) carta(s) específica(s) preenchida(s) e com as devidas assinaturas (0628-A para apólices do Saúde 
	e 0628-B para apólices do Dental).</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" align="left" style="border-bottom:1px solid #000000">
	<b> 3 - Endereço residencial do titular (para correspondência)</b>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="56%">
	Logradouro (Rua, Avenida, Praça, etc) nº e complemento (andar, sala, apto, etc)<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="24%">
	Bairro<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="6%">
	UF<br>&nbsp;		</td>
	<td class="fundo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" width="14%">
	&nbsp;<br>&nbsp;		</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000" width="37%">
	Cidade<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000" width="13%">
	CEP<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:1px solid #000000" width="14%">
	Telefone (com DDD)<br>&nbsp;		</td>
	<td class="fundo" valign="top" align="left" style="border-bottom:2px solid #000000;border-right:0px solid #000000" width="36%">
	&nbsp;<br>&nbsp;		</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" align="left" style="border-bottom:1px solid #000000">
	<b> 4 - Dados do plano</b>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="11%">
	Código do plano:<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="10%">
	Código da região:<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="29%">
	Nome da região:<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="14%">
	Data da alteração<br>&nbsp;		</td>
	<td class="fundo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" width="36%">
	&nbsp;<br>&nbsp;		</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" align="left" style="border-top:2px solid #000000">	<b> JUNHO/2012</b>	</td>
	<td class="campo" align="center" style="border-top:2px solid #000000">	<b> PAG.1/3</b>	</td>
	<td class="campo" align="right" style="border-top:2px solid #000000">	<b>COD. FORM. ELETR. 0628 </b>	</td>
</tr>
</table>

</center></div>
<!-- ----------------------------- -->
<DIV style="page-break-after:always"></DIV>

<div align="center"><center>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" align="left" width="25%" style="border-bottom:1px solid #000000"><img src="../images/bradesco_saude.jpg" border="0"></td>
	<td class="campop" align="left" valign="top" width="50%" style="border-bottom:1px solid #000000"><b>
	Bradesco Saúde Coletivo Empresarial/SPG<br>
	Formulário de Cancelamento/Alteração de Dados do Segurado e Dependente</b>
	</td>
	<td class="campo" align="left" width="25%" style="border-bottom:1px solid #000000"><img src="../images/bradesco_dental.jpg" border="0"></td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Companhia Seguradora<br>Bradesco Saúde S.A.</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	CNPJ<br>92.693.118/0001-60</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Registro na ANS:<br>005711</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Operadora<br>Odontoprev S.A.</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	CNPJ<br>58.119.199/0001-51</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Registro na ANS:<br>30194-9</td>
</tr>
<tr>
	<td class="campo" align="left" colspan="6" style="border-top:2px solid #000000;border-bottom:1px solid #000000">
	<b> 5 - Dependentes</b>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Favor preencher os devidos campos com<br>os códigos informados ao lado<br>&nbsp;		</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Sexo: <br> 1 - Masculino <br> 2 - Feminino</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Estado Civil: <br> 1 - Solteiro <br> 2 - Casado	</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	&nbsp; <br> 3 - Viúvo <br> 4 - Outros	</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Parentesco: <br> 1 - Cônjuge <br> 2 - Filho(a)</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	&nbsp; <br> 3 - Mãe <br> 4 - Pai</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	&nbsp; <br> 5 - Sogro <br> 6 - Sogra</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	&nbsp; <br> 7 - Tutelado <br> 8 - Outros</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Motivo do Cancelamento: <br> 1 - Desistência <br> 6 - Portabilidade</td>
	<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	&nbsp; <br> 4 - Falecimento &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 5 - Duplicidade	</td>
</tr>
</table>

<%for a=1 to 3%>
	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
	<tr>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="9%">
		Certificado<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="10%">
		Código<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="50%">
		Nome<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="16%">
		CPF<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" width="15%">
		Data de nascimento<br>&nbsp;		</td>
	</tr>
	</table>

	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
	<tr>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="4%">
		Sexo<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="5%">
		Est.Civil<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="10%">
		Parentesco<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="50%">
		Nome da Mãe<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="25%">
		Data de inclusão ou alteração ou cancelamento<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" width="6%">
		Motivo<br>&nbsp;		</td>
	</tr>
	</table>

	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
	<tr>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="34%">
		CNS (Carteira Nacional de Saúde):<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="33%">
		DNV (Declaração de Nascido Vivo):<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" width="33%">
		RIC (Registro de Identidade Civil):<br>&nbsp;		</td>
	</tr>
	</table>

	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
	<tr>
		<td class="fundo" style="border-bottom:1px solid #000000" height="10px"></td>
	</tr>
	</table>
<%next%>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" align="left" style="border-top:2px solid #000000">	<b> JUNHO/2012</b>	</td>
	<td class="campo" align="center" style="border-top:2px solid #000000">	<b> PAG.2/3</b>	</td>
	<td class="campo" align="right" style="border-top:2px solid #000000">	<b>COD. FORM. ELETR. 0628 </b>	</td>
</tr>
</table>

</center></div>

<!-- ----------------------------- -->
<DIV style="page-break-after:always"></DIV>

<div align="center"><center>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" align="left" width="25%" style="border-bottom:1px solid #000000"><img src="../images/bradesco_saude.jpg" border="0"></td>
	<td class="campop" align="left" valign="top" width="50%" style="border-bottom:1px solid #000000"><b>
	Bradesco Saúde Coletivo Empresarial/SPG<br>
	Formulário de Cancelamento/Alteração de Dados do Segurado e Dependente</b>
	</td>
	<td class="campo" align="left" width="25%" style="border-bottom:1px solid #000000"><img src="../images/bradesco_dental.jpg" border="0"></td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Companhia Seguradora<br>Bradesco Saúde S.A.</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	CNPJ<br>92.693.118/0001-60</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Registro na ANS:<br>005711</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	Operadora<br>Odontoprev S.A.</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	CNPJ<br>58.119.199/0001-51</td>
	<td class="campo" valign="top" align="center" style="border-bottom:1px solid #000000;border-right:0px solid #000000">
	Registro na ANS:<br>30194-9</td>
</tr>
<tr>
	<td class="campo" align="left" colspan="6" style="border-top:2px solid #000000;border-bottom:1px solid #000000">
	<b> 5 - Dependentes (continuação)</b>
	</td>
</tr>
</table>

<%for a=1 to 2%>
	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
	<tr>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="9%">
		Certificado<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="10%">
		Código<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="50%">
		Nome<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="16%">
		CPF<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" width="15%">
		Data de nascimento<br>&nbsp;		</td>
	</tr>
	</table>

	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
	<tr>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="4%">
		Sexo<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="5%">
		Est.Civil<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="10%">
		Parentesco<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="50%">
		Nome da Mãe<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="25%">
		Data de inclusão ou alteração ou cancelamento<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" width="6%">
		Motivo<br>&nbsp;		</td>
	</tr>
	</table>

	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
	<tr>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="34%">
		CNS (Carteira Nacional de Saúde):<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:1px solid #000000" width="33%">
		DNV (Declaração de Nascido Vivo):<br>&nbsp;		</td>
		<td class="campo" valign="top" align="left" style="border-bottom:1px solid #000000;border-right:0px solid #000000" width="33%">
		RIC (Registro de Identidade Civil):<br>&nbsp;		</td>
	</tr>
	</table>

	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
	<tr>
		<td class="fundo" style="border-bottom:1px solid #000000" height="10px"></td>
	</tr>
	</table>
<%next%>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" align="left" colspan="1" style="border-top:2px solid #000000;border-bottom:0px solid #000000">
	Declaro a veracidade das informações prestadas neste formulário e autorizo o seu processamento no sistema da Bradesco Saúde.
	</td>
</tr>
<tr>
	<td class="campo" align="center" colspan="1" style="border-top:0px solid #000000;border-bottom:0px solid #000000">
	<br>
	<br>
	<br>
	<br>___________________________________________________________________
	<br>Assinatura do Estipulante sob carimbo
	<br>
	<br>
	<br>
	<br>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="1090">
<tr>
	<td class="campo" align="left" style="border-top:2px solid #000000">	<b> JUNHO/2012</b>	</td>
	<td class="campo" align="center" style="border-top:2px solid #000000">	<b> PAG.3/3</b>	</td>
	<td class="campo" align="right" style="border-top:2px solid #000000">	<b>COD. FORM. ELETR. 0628 </b>	</td>
</tr>
</table>

</center></div>

<!-- ----------------------------- -->
<DIV style="page-break-after:always"></DIV>

	
<%for a=1 to 10:response.write "<br>":next%>







<%
elseif temp=2 then
%>
<table border="1" cellpadding="0" width="550" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>&nbsp;Chapa</td>
	<td class=titulo>&nbsp;Nome</td>
	<td class=titulo>&nbsp;Situacao</td>
</tr>
<%
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo>&nbsp;<%=rs("chapa")%></td>
	<td class=campo>&nbsp;<a href="decl_copart_bradesco.asp?codigo=<%=rs("chapa")%>"><%=rs("nome")%></a></td>
	<td class=campo>&nbsp;<%=rs("codsituacao")%></td>
</tr>
<%
rs.movenext
loop
%>

</table>
<%
rs.close
set rs=nothing
end if ' temps

conexao.close
set conexao=nothing
set rs3=nothing
conexao2.close
set conexao2=nothing
%>
</body>
</html>