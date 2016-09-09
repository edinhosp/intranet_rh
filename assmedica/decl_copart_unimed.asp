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
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para Declaração opcional de Plano de Saúde - UNIMED
<form method="POST" action="decl_copart_unimed.asp">
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
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td valign="center" rowspan=1 height=21 align="left" width=100% valign=top style="border-bottom:2px solid #000000"></td>
	<td rowspan=2 align="right" valign=top width=146 height=48><img src="../images/logo_unimed.jpg" border="0"></td>
</tr>
<tr>
	<td valign="center" rowspan=1 height=27 align="left" valign=top><font style="font-size:14px"><b>Cadastro Seguros Saúde<br>Inativos</b></font></td>
</tr>
<tr><td colspan=2 class=campo height=20></td></tr>
<tr><td colspan=2 class="campop" align="left" height=20 style="border-bottom:2px solid #000000"><b>Dados do Titular</td></tr>
</table>
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
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:0 solid #000000" width="690">
<tr><td valign=top>

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
		<tr><td class="campor">Nome do Titular</td>
			<td class="campor" width=250 style="border-left:1px solid">Código do Cartão de identificação (quando ativo)</td>
		</tr>
		<tr><td class=campo style="border-bottom: 1px solid">&nbsp;<%=rs("nome")%></td>
			<td class=campo style="border-bottom: 1px solid;border-left:1px solid">&nbsp;<%=carteirinha%></td>
		</tr>
	</table>

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
		<tr><td class="campor" width=18%>Nome da Mãe</td></tr>
		<tr><td class=campo style="border-bottom: 1px solid">&nbsp;<%=mae%></td></tr>
	</table>

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
		<tr>
			<td class="campor" style="border-right:1px solid">CPF</td>
			<td class="campor" colspan=5 style="border-right:0px solid">Situação</td>
		</tr>
		<tr>
			<td style="border-bottom:2px solid;border-right:1px solid" class=campo>&nbsp;<%=rs("cpf")%></td>
			<td style="border-bottom:2 solid"><img src="../images/round_square<%=bolas%>.jpg" border="0"></td><td class=campo style="border-bottom:2 solid">Aposentado</td>
			<td style="border-bottom:2 solid"><img src="../images/round_square<%=bolac%>.jpg" border="0"></td><td class=campo style="border-bottom:2 solid">Demitido sem justa causa</td>		
			<td class=campo width=30% style="border-bottom:2px solid"></td>	
		</tr>
	</table>
	
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td colspan=2 class="campop" align="left" height=20 style="border-bottom:2px solid #000000"><b>Dados do Dependente</td></tr>
</table>

<%
sqld="SELECT d.chapa, d.dependente, d.nascimento, d.sexo, d.parentesco, d.cpf, d.mae, p.empresa, p.plano " & _
"FROM assmed_dep d, assmed_dep_mudanca p WHERE d.chapa=p.chapa and d.nrodepend=p.nrodepend and p.plano='" & plano & "' " & _
"AND d.chapa='" & rs("chapa") & "' AND '" & dtaccess(rsdatademissao) & "' Between p.ivigencia And p.fvigencia "
rs3.Open sqld, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
%>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:0 solid #000000" width="690">
<tr>
	<td valign=middle align="center" style="border-left:0 solid;border-bottom: 1px solid;border-right: 1px solid;background-color:white;color:black" width=20><b><%=rs3.absoluteposition%></td>
	<td>
	<!-- -->
	<table border=0 bordercolor=#000000 width=100% style="border-collapse:collapse">
	<tr><td class="campor" style="border-right:0 solid">Nome </td></tr>
	<tr><td style="border-bottom: 1px solid;border-right:0 solid" class="campor">&nbsp;<%=rs3("dependente")%></td></tr>
	<tr><td class="campor" style="border-right:0 solid">Nome da Mãe do Dependente </td></tr>
	<tr><td style="border-bottom: 1px solid;border-right:0 solid" class="campor">&nbsp;<%=rs3("mae")%></td></tr>
	</table>	
<!-- -->
	</td>
</tr></table>
<%
rs3.movenext:loop
end if 'rs3.recordcount>0

if rs3.recordcount=0 or rs3.recordcount<4 then
for b=rs3.recordcount+1 to 4
%>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:0 solid #000000" width="690">
<tr><td valign=middle align="center" style="border-left:0 solid;border-bottom: 1px solid;border-right: 1px solid;background-color:white;color:black" width=20><b><%=b%></td>
	<td>
	<table border=0 bordercolor=#000000 width=100% style="border-collapse:collapse">
	<tr><td class="campor" style="border-right:0 solid">Nome </td></tr>
	<tr><td style="border-bottom: 1px solid;border-right:0 solid" class="campor">&nbsp;</td></tr>
	</table>	

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
	<tr><td class="campor" style="border-right:0 solid">Nome da Mãe do Dependente </td></tr>
	<tr><td style="border-bottom: 1px solid;border-right:0 solid" class="campor">&nbsp;</td></tr>
	</table>	
<!-- -->
	</td>
</tr></table>
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

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" height=22 style="border-bottom:1px solid"><b>Tempo em que o Segurado contribuiu para o Plano</td>
	<td class=campo width=200 style="border-bottom:1px solid;border-left:1px solid">&nbsp;<%=cano%> anos <%=cmes%> meses</td>
</tr>
<tr>
	<td class="campor" height=22 style="border-bottom:1px solid"><b>Valor do último desconto a título de contribuição no seguro saúde</td>
	<td class=campo width=200 style="border-bottom:1px solid;border-left:1px solid">&nbsp;R$ <%=formatnumber(ultimo,2)%></td>
</tr>
<tr>
	<td class="campor" height=22 style="border-bottom:1px solid"><b>Data de desligamento da empresa</td>
	<td class=campo width=200 style="border-bottom:1px solid;border-left:1px solid">&nbsp;<%=dia3%>&nbsp;/&nbsp;<%=mes3%>&nbsp;/&nbsp;<%=ano3%></td>
</tr>
<tr>
	<td class="campor" height=22 style="border-bottom:1px solid"><b>Data de término do benefício/acordo coletivo (se houver)</td>
	<td class=campo width=200 style="border-bottom:1px solid;border-left:1px solid">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;/&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" height=22 style="border-bottom:1px solid"><b>Início/Término da Vigência no Plano (Uso exclusivo Seguradora)</td>
	<td class=campo width=200 style="border-bottom:1px solid;border-left:1px solid">&nbsp;Início &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Término &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;/</td>
</tr>

<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
<tr>
	<td class="campor" colspan=6 style="border-right:0px solid">Plano</td>
</tr>
<tr>
	<td class=campo width=5% style="border-bottom:2px solid"></td>	
	<td style="border-bottom:2 solid"><img src="../images/round_square<%=bolas%>.jpg" border="0"></td><td class=campo style="border-bottom:2 solid">Permanecer no mesmo Plano</td>
	<td style="border-bottom:2 solid"><img src="../images/round_square<%=bolac%>.jpg" border="0"></td><td class=campo style="border-bottom:2 solid">Reduzir para o Plano</td>		
	<td class=campo width=30% style="border-bottom:2px solid"></td>	
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td colspan=2 class="campop" align="left" height=20 style="border-bottom:2px solid #000000"><b>Dados para cobrança</td></tr>
</table>

<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
<tr><td class="campor" style="border-right: 1px solid">Endereço</td>
	<td class="campor" style="border-right:0 solid">Complemento</td>
</tr>
<tr><td style="border-bottom: 1px solid;border-right: 1px solid" width=60% class=campo>&nbsp;<%=rs("rua")%>&nbsp;<%=rs("numero")%></td>
	<td style="border-bottom: 1px solid;border-right:0 solid" class=campo>&nbsp;<%if isnull(rs("complemento")) or rs("complemento")="" then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" else response.write rs("complemento")%></td>
</tr>
</table>

<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
<tr><td class="campor" style="border-right: 1px solid">Bairro</td>
	<td class="campor" style="border-right: 1px solid">Cidade</td>
	<td class="campor" style="border-right:0 solid">UF</td>
</tr>
<tr>
	<td style="border-bottom: 1px solid;border-right: 1px solid" class=campo>&nbsp;<%if isnull(rs("bairro")) or rs("bairro")="" then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" else response.write rs("bairro")%></td>
	<td style="border-bottom: 1px solid;border-right: 1px solid" class=campo>&nbsp;<%=rs("cidade")%></td>
	<td style="border-bottom: 1px solid;border-right:0 solid" class=campo>&nbsp;<%=rs("estado")%></td>
</tr></table>

<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
<tr><td class="campor" width=15% style="border-right: 1px solid">CEP</td>
	<td class="campor" width=25% style="border-right: 1px solid">Telefone</td>
	<td class="campor" width=25% style="border-right: 1px solid">Fax</td>
	<td class="campor">E-mail</td>
</tr>
<tr><td style="border-bottom:2 solid;border-right: 1px solid" class=campo>&nbsp;<%=rs("cep")%></td>
	<td style="border-bottom:2 solid;border-right: 1px solid" class=campo>&nbsp;<%=rs("telefone1")%></td>
	<td style="border-bottom:2 solid;border-right: 1px solid" class=campo>&nbsp;</td>
	<td style="border-bottom:2 solid" class=campo>&nbsp;<%=rs("email")%></td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td colspan=2 class="campop" align="left" height=20 style="border-bottom:2px solid #000000"><b>Observações</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class=campo>
O funcionário está ciente de que:
<ol type="a" style="margin-top:0px;margin-bottom:0px">
	<li>a cobertura tem validade a partir do 1º dia do mês subsequente à solicitação de adesão e pelo tempo estabelecido pela Lei nº 9656/98;</li>
	<li>o direito de permanência cessará, definitivamente, na data em que o ex-empregado assumir um novo emprego ou atividade remunerada, 
	<b>responsabilizando-se em comunicar por escrito à Unimed Seguros Saúde a sua admissão sob pena de, não o fazendo, ser responsabilizado 
	pessoalmente por todas as despesas assumidas pela Seguradora durante o período não comunicado, além das perdas e danos;</b></li>
	<li>a permanência no seguro saúde será de acordo com o mesmo plano que tinha enquanto empregado ou plano inferior, caso a empresa tenha 
	contratado para o grupo ativo e, desde que o empregado faça essa opção por seu <b>exclusivo interesse;</b></li>
	<li><b>será responsável pelo pagamento integral do prêmio</b>;</li>
	<li>estarão inclusos todos os dependentes que já constavam no plano anteriormente, exceto os que perderem a condição de dependência; </li>
	<li>o Segurado é responsável pela declaração de dependentes, <b>respondendo por qualquer falha ou omissão que possa induzir a Seguradora a erro 
	quanto à aceitação daqueles</b>;</li>
	<li>a Seguradora reserva-se o direito de exigir documentação comprobatória de dependência econômica, sempre que julgar necessário.</li>
	<li>caso seja rescindido o contrato do estipulante ao qual o segurado principal estava vinculado quando ativo, seu contrato estará, automaticamente, cancelado;</li>
	<li>o Plano de Inativos contempla somente as coberturas médico-hospitalares e não inclui os Benefícios Especiais, tais como SEA e outros.</li>
</ol>
</td></tr>
</table>

<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
<tr><td class=campo width=15% style="border-right:0 solid;border-top:1px solid">
	Informar abaixo quais os Segurados <b>DEPENDENTES</b> que <b>NÃO</b> deverão permanecer no plano e justificar a perda da dependência.</td>
</tr>
<tr><td class=campo height=20 style="border-bottom:1px solid"</td></tr>
<tr><td class=campo height=20 style="border-bottom:1px solid"</td></tr>
<tr><td class=campo height=20 style="border-bottom:1px solid"</td></tr>
<tr><td class=campo height=20 style="border-bottom:1px solid"</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td colspan=2 class="campop" align="left" height=20 style="border-bottom:2px solid #000000"><b>Assinaturas</td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" height=40 style="border-right: 1px solid #000000;border-bottom:2px solid" valign=top>Assinatura da Empresa (sob carimbo contendo o CNPJ/MF)<td>
	<td class="campor" width=50% style="border-bottom:2px solid #000000" valign=top>Assinatura do Segurado<td>
</tr>
</table>	

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo valign=top align="left"><img src="../images/ans_unimed.gif" border="0" alt=""></td>
	<td class=campo align="right">
	<b>Unimed Seguros Saúde S.A.</b> - CNPJ/MF 04.487.255/0001-81<br>
	Alameda Santos, 1827 4º andar CEP 01419-909 São Paulo SP<br>
	Tel: Grande São Paulo 3265-9672 - Demais localidades 0800 16 6633<br>
	www.unimedseguros.com.br
	</td>
</tr>
</table>


<%
rs.close
set rs=nothing

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
	<td class=campo>&nbsp;<a href="decl_copart.asp?codigo=<%=rs("chapa")%>"><%=rs("nome")%></a></td>
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