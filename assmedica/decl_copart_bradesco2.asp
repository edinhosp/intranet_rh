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
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para Declaração opcional de Plano de Saúde (Carta) - BRADESCO
<form method="POST" action="decl_copart_bradesco2.asp">
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

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" align="left" width="66%" style="border-bottom:0px solid #000000"><img src="../images/aguia_color.jpg" border="0" height="95px">
	<br>Osasco, <%=formatdatetime(now(),1)%>
	<br>Sr<%if sexo="F" then response.write "a"%> <%=nome%>
	<br>Data do Desligamento: <%=rsdatademissao%>
	
	</td>
	<td class="campo" align="center" valign="top" width="34%" style="border:1px solid #000000">
	Data de Protocolo do Recebimento<br>
	da Carta pelo Ex-Empregado:<br>
	<br>
	______/_______/_______<br>
	<br>
	<br>
	______________________________<br>
	<i>(Assinatura do Ex-Empregado/Titular)</i>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" align="left" style="border:0px solid #000000">
	<br><br><br><br><br><br>
	<div align="center"><b>SEGURO - SAÚDE</b></div>
	<br><br>
	Tendo em vista seu desligamento desta Empresa na data acima identificada, e de acordo com o disposto na <b>Lei nº 9.656/1998</b> e 
	na <b>Resolução Normativa RN nº 279/2011</b> e suas atualizações publicadas pela Agência Nacional de Saúde Suplementar - ANS, valemo-nos 
	da presente para oferecer-lhe a <u>opção de permanência no seguro-saúde contratado com a BRADESCO SAÚDE</u>, nas mesmas condições de 
	cobertura assistencial de que gozava quando da vigência do contrato de trabalho, pelo prazo de vigência e condições abaixo especificadas, 
	obedecidas as condições do Contrato de Seguro firmado entre esta empresa e a BRADESCO SAÚDE.<br>
	<br><br><br>
	<b>Condição: (&nbsp;&nbsp;&nbsp;) Aposentado (&nbsp;&nbsp;&nbsp;) Demitidos sem justa causa (&nbsp;&nbsp;&nbsp;) Aposentado que continuou trabalhando</b>
	<br><br><br>
	Prazo de Vigência: ________ (número de meses), a partir de ___/___/____, (data do início do benefício), observado o disposto no item 8 da 
	Declaração que integra esta correspondência.
	<br><br><br>
	Sendo assim, solicitamos que formalize a sua opção, preenchendo o(s) quadro(s) adiante e devolva uma via desta, devidamente datada e assinada 
	para esta empresa em até 30 (trinta) dias a contar da data de recebimento desta carta, impreterivelmente.
	<br><br><br><br>
	Atenciosamente,
	<br><br><br>
	__________________________________<br>
	<i>(Assinatura e Carimbo da Empresa)</i>
	<br><br><br><br><br><br><br><br><br>
	</td>
</tr>
</table>

<br><br><br><br>


<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campo" align="left" style="border-top:2px solid #000000">	<b> JUNHO/2012</b>	</td>
	<td class="campo" align="right" style="border-top:2px solid #000000">	<b>FORM. ELETR. 0628-A </b>	</td>
</tr>
</table>

</center></div>
<!-- ----------------------------- -->
<DIV style="page-break-after:always"></DIV>

<div align="center"><center>
<br>
<br>
<br>
<table border="0" cellpadding="3" cellspacing="2" style="border-collapse: collapse" width="690">
<tr>
	<td class="campo" align="left" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<b> EM RELAÇÃO AO SEGURO-SAÚDE CONTRATADO COM A BRADESCO SAÚDE E TENDO CONHECIMENTO DAS CONDIÇÕES E
	COBERTURAS PARA DEMITIDOS SEM JUSTA CAUSA E APOSENTADOS, DECLARO A MINHA OPÇÃO DE:</b>
	</td>
</tr>
<tr>
	<td class="campo" align="left" style="border-left:1px solid #000000;border-right:1px solid #000000" height="25px">
	<b>(&nbsp;&nbsp;&nbsp;) PERMANECER no seguro-saúde, como titular, e MANTER os dependentes listados abaixo:</b>
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
		<tr><td class="campo" height="20px">Dependente 1:</td><td class="campo" colspan="5" style="border-bottom:1px solid #000000"></td></tr>
		<tr><td class="campo" height="20px">Data Nasc.:</td><td class="campo">_____/_____/________</td>
		<td class="campo">CPF nº</td><td class="campo" style="border-bottom:1px solid #000000" width="19%"></td>
		<td class="campo">Nome da Mãe do Depen.1:</td><td class="campo" style="border-bottom:1px solid #000000" width="29%"></td></tr>
		</table>
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
		<tr><td class="campo" height="20px">Dependente 2:</td><td class="campo" colspan="5" style="border-bottom:1px solid #000000"></td></tr>
		<tr><td class="campo" height="20px">Data Nasc.:</td><td class="campo">_____/_____/________</td>
		<td class="campo">CPF nº</td><td class="campo" style="border-bottom:1px solid #000000" width="19%"></td>
		<td class="campo">Nome da Mãe do Depen.2:</td><td class="campo" style="border-bottom:1px solid #000000" width="29%"></td></tr>
		</table>
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
		<tr><td class="campo" height="20px">Dependente 3:</td><td class="campo" colspan="5" style="border-bottom:1px solid #000000"></td></tr>
		<tr><td class="campo" height="20px">Data Nasc.:</td><td class="campo">_____/_____/________</td>
		<td class="campo">CPF nº</td><td class="campo" style="border-bottom:1px solid #000000" width="19%"></td>
		<td class="campo">Nome da Mãe do Depen.2:</td><td class="campo" style="border-bottom:1px solid #000000" width="29%"></td></tr>
		</table>
	</td>
</tr>
<tr>
	<td class="campo" align="left" style="border-left:1px solid #000000;border-right:1px solid #000000" height="25px">
	<b>(&nbsp;&nbsp;&nbsp;) PERMANECER no seguro-saúde, como titular, e NÃO MANTER dependentes.</b>
	</td>
</tr>
<tr>
	<td class="campo" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000" height="25px">
	<b>(&nbsp;&nbsp;&nbsp;) SER EXCLUÍDO do seguro-saúde, juntamente com todo meu grupo familiar de forma irrevogável e irretratável.</b>
	</td>
</tr>
</table>
<br>
<br>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" align="left" style="border:0px solid #000000" colspan="2">
	<br><br><br>
	<div align="center"><b>DECLARAÇÃO</b></div>
	<br><br><br>
	<b>Declaro ter ciência e concordar com as condições abaixo:</b>
	<br><br><br>
	<ol>
	<li>Sou responsável, a partir da data de início de vigência acima especificado, pelo <b>pagamento integral do prêmio de meu seguro-saúde e de meus
		dependentes listados acima</b>;<br>
	<li>Deverei <b>manter atualizado</b>, junto à empresa os meus <b>dados cadastrais</b> (endereço completo, fone);<br>
	<li>Não poderei, durante o tempo de vigência, <b>mudar de padrão de plano</b>;<br>
	<li><b>A cobertura do seguro-saúde será suspensa em caso de atraso no pagamento por período superior a 30 (trinta) dias</b>;<br>
	<li><b>O seguro-saúde será cancelado, de forma definitiva e irreversível, quando uma das mensalidades permanecer pendente de pagamento pelo prazo
		de 60 (sessenta) dias, consecutivos ou não, no período de 12 (doze) meses</b>;<br>
	<li>A minha relação com o Seguro Saúde Bradesco está condicionada à manutenção da apólice para os funcionários ativos;<br>
	<li>Tenho conhecimento do valor integral do prêmio de meu seguro-saúde, bem como da forma e periodicidade de reajuste desse valor, que terá como
		data-base o mês de <u>&nbsp;&nbsp;Outubro&nbsp;&nbsp;</u> <i>(mês de reajuste do contrato)</i> em conformidade com as cláusulas de reajustes previstas
		no contrato;<br>
	<li><b>Meu direito de manutenção no seguro saúde, bem como o de meus dependentes, se extingue após decorrido o prazo de vigência especificado acima, ou antes
		disso caso ocorra uma das seguintes hipóteses:<br>
	<ul style="list-style-type:lower-roman">
		<li>Quando da minha admissão em novo emprego, considerando-se qualquer novo vínculo profissional que possibilite o meu ingresso em um plano de assistência a
		saúde coletivo empresarial, coletivo por adesão ou de autogestão.<br>
		<li>Quando do cancelamento, pela empresa, do seguro de assistência à saúde oferecido aos seus empregados ativos e ex-empregados.
		<li>Por inexatidão ou omissão no preenchimento do documento de inclusão, que tenha influenciado na aceitação do seguro, mediante apresentação de prova pela
		Seguradora e comunicação escrita ao Estipulante.
		<li>Em caso de infrações ou frades compravadas; ou
		<li>Por minha solicitação formal ao Estipulante.</br></b>
	</ul>
	</ol>
	</td>
</tr>
</table>

<br><br>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campo" align="left" style="border-top:2px solid #000000">	<b> JUNHO/2012</b>	</td>
	<td class="campo" align="right" style="border-top:2px solid #000000">	<b>FORM. ELETR. 0628-A </b>	</td>
</tr>
</table>

</center></div>
<!-- ----------------------------- -->
<DIV style="page-break-after:always"></DIV>

<div align="center"><center>
<br>
<br>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" align="left" style="border:0px solid #000000" colspan="2">

	<ol start="9">
	<li>É de minha responsabilidade informar à empresa tão logo eu seja admitido em novo emprego ou qualquer outro vínculo profissional que possibilite o meu ingresso
		em um plano de assistência à saúde coletivo empresarial, coletivo por adesão ou de autogestão.<br>
	<li>Recebi, tomei conhecimento e concordo com as condições estabelecidas no Manual do Segurado.<br>
	</ol>
	</td>
</tr>
<tr>
	<td class="campop" align="left" style="border:0px solid #000000">
	<br><br><br><br><br>	<br><br><br><br><br>

	_____________________________________________<br>
	<i>(Local e Data)</i>
	</td>
	<td class="campop" align="left" style="border:0px solid #000000">
	<br><br><br><br><br>	<br><br><br><br><br>

	_____________________________________________<br>
	<i>(Assinatura do Ex-empregado/segurado titular)</i>

	</td>
</tr>
</table>
	
<%
for a=1 to 40
response.write "</br>"
next
%>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campo" align="left" style="border-top:2px solid #000000">	<b> JUNHO/2012</b>	</td>
	<td class="campo" align="right" style="border-top:2px solid #000000">	<b>FORM. ELETR. 0628-A </b>	</td>
</tr>
</table>

</center></div>


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