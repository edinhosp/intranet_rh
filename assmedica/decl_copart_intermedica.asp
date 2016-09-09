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

set rs=server.createobject ("ADODB.Recordset")
set rs2=server.createobject ("ADODB.Recordset")
set rs3=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
Set rs2.ActiveConnection = conexao
Set rs3.ActiveConnection = conexao

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

	sqla="SELECT f.NOME, f.CODSITUACAO, f.CHAPA, 'DataAdmissao'=f.ADMISSAO, f.CODSECAO, f.codsindicato, f.Secao, f.dtnascimento, " & _
	"f.telefone1, f.telefone2, f.telefone3, f.email, f.cpf, f.estadocivil, f.funcao, f.cartidentidade, f.cpf, f.dtnascimento, " & _
	"f.sexo, f.rua, f.numero, f.complemento, f.bairro, f.cidade, f.cep, f.estado, 'DataDemissao'=f.demissao, f.dtaposentadoria, " & _
	"f.aposentado, f.tipodemissao, f.grauinstrucao, f.pispasep, i.CARTAOSUS, f.mae " & _
	"FROM qry_funcionarios f left join corporerm.dbo.VPCOMPL i on i.CODPESSOA=f.CODPESSOA " & _
	"where f.codpessoa>0 "
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
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para Declaração opcional de Plano de Saúde - INTERMéDICA
<form method="POST" action="decl_copart_intermedica.asp">
	<p style="margin-top: 0; margin-bottom: 0">Chapa/Nome <input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>">
	<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>

<%
elseif temp=0 then
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
idade=int((now()-rs("dtnascimento"))/365.25)
if rs("datademissao")="" or isnull(rs("datademissao")) then rsdatademissao=now() else rsdatademissao=rs("datademissao")

sqlplano="SELECT codigo, plano FROM assmed_mudanca " & _
"WHERE chapa='" & rs("chapa") & "' and '" & dtaccess(rsdatademissao) & "' between ivigencia and fvigencia "
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly
plano=rs3("plano")
carteirinha=rs3("codigo")
rs3.close

if rs("aposentado")=1 or rs("aposentado")=true then 
	Sapos="&nbsp;X&nbsp;":Napos="&nbsp;&nbsp;&nbsp;&nbsp;"
else 
	Sapos="&nbsp;&nbsp;&nbsp;&nbsp;":Napos="&nbsp;X&nbsp;"
end if
if rs("tipodemissao")="2" or rs("tipodemissao")="A" then 
	Sdem="&nbsp;X&nbsp;":Ndem="&nbsp;&nbsp;&nbsp;&nbsp;"
else 
	Sdem="&nbsp;&nbsp;&nbsp;&nbsp;":Ndem="&nbsp;X&nbsp;"
end if

idade=int((now()-rs("dtnascimento"))/365.25)
dia4=day(rs("dtaposentadoria")):if dia4="" or isnull(dia4) then dia4="  " else dia4=numzero(dia4,2)
mes4=month(rs("dtaposentadoria")):if mes4="" or isnull(mes4) then mes4="  " else mes4=numzero(mes4,2)
ano4=year(rs("dtaposentadoria")):if ano4="" or isnull(ano4) then ano4="  " else ano4=right(ano4,2)
dtaposent=dia4&mes4&ano4

'052 desconto co-participação 076 desconto assistencia médica
sqlp="select vezes=COUNT(chapa), primeira=MIN(periodo), ultima=MAX(periodo) from ( " & _
"select distinct chapa, periodo=convert(nvarchar(4),ANOCOMP)+case when mescomp<10 then '0' else '' end+CONVERT(nvarchar(2),mescomp) from corporerm.dbo.pffinanc where codevento IN ('076','076S','076I','076U','052','052I') " & _
"UNION " & _
"select distinct chapa, periodo=convert(nvarchar(4),ANOCOMP)+case when mescomp<10 then '0' else '' end+CONVERT(nvarchar(2),mescomp) from corporerm.dbo.pffinanccompl where codevento IN ('076','076S','076I','076U','052','052I') " & _
") z where CHAPA='" & rs("chapa") & "' " 
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
mes1=rs3("vezes")
if mes1>0 then
	Scont="&nbsp;X&nbsp;":Ncont="&nbsp;&nbsp;&nbsp;&nbsp;"
else 
	Scont="&nbsp;&nbsp;&nbsp;&nbsp;":Ncont="&nbsp;X&nbsp;"
end if
contrib1=rs3("primeira")
contrib99=rs3("ultima")
if mes1="" or isnull(mes1) then mes1=0
cano=int((mes1)/12)
cmes=(mes1)-(cano*12)
dini=dtdemissao
rs3.close

sqlp="select max(valor) ultima from corporerm.dbo.pffinanc where codevento in ('052','052U','052I','076','076I','076U') and chapa='" & rs("chapa") & "' " & _
" and dtpagto>=dateadd(m,-2,'" & dtaccess(rs("datademissao")) & "') "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
ultimo=rs3("ultima")
if ultimo="" or isnull(ultimo) then ultimo=0
rs3.close

'response.write "<br>" & mes1
'response.write "<br>" & cano
'response.write "<br>" & cmes
'response.write "<br>" & ultimo
'response.write "<br>" & contrib1
'response.write "<br>" & contrib99
%>

<div align="center">
<center>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campo" height="62px" width="204px" style="border-right:0px"><img src="../images/intermedica_notredame_ans.png" height="60px" width="202px" border="0" alt=""></td>
	<td class="campop" valign="middle" align="center" style="border-left:0px;border-right:0px" ><span style="font-size:14pt"><b><i>Termo de Opção de Continuidade</i></b></span><br>(Conforme Arts. 30 e 31 da Lei 9.656/79 e RN 279)</td>
	<td class="campor" style="border-left:0px">[&nbsp;&nbsp;&nbsp;] DEMITIDO ou EXONERADO<br><br>[&nbsp;&nbsp;&nbsp;] APOSENTADO	</td>
</tr>
<tr>
	<td colspan="3" class="campor" style="background-color:black;color:white"><b>&nbsp;1 - DADOS CADASTRAIS TITULAR</b></td>
</tr>
</table>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor"><font size="1">Cód.Empresa atual:</font>	<br>&nbsp;<b>18980000</td>
	<td class="campor"><font size="1">Matrícula Vigente</font>	<br>&nbsp;<b><%=rs("chapa")%> </td>
	<td class="campor"><font size="1">Plano Vigente:</font>	<br>&nbsp;<b><%=plano%> </td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nome Completo:</font>	<br>&nbsp;<b><%=rs("nome")%> </td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Data de Nascimento:</font>	
	<br>&nbsp;<b><%=day(rs("dtnascimento")) & " | " & month(rs("dtnascimento")) & " | " & year(rs("dtnascimento")) %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Sexo:</font>	
	<br>&nbsp;<b><%=rs("sexo") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">E C</font>	
	<br>&nbsp;<b><%=rs("estadocivil") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">CPF</font>	
	<br>&nbsp;<b><%=rs("cpf") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">PIS/PASEP</font>	
	<br>&nbsp;<b><%=rs("pispasep") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Cartão Nacional de Saúde</font>	
	<br>&nbsp;<b><%=rs("cartaosus") %>  </td>
</tr>
</table>
	
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nome Completo da Mãe:</font>	<br>&nbsp;<b><%=rs("mae")%> </td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">CEP</font>	
	<br>&nbsp;<b><%=rs("cep") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Rua / Av.</font>	
	<br>&nbsp;<b><%=rs("rua") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nº</font>	
	<br>&nbsp;<b><%=rs("numero") %>  </td>
</tr>
</table>	

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Bairro</font>	
	<br>&nbsp;<b><%=rs("bairro") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Cidade</font>	
	<br>&nbsp;<b><%=rs("cidade") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">UF</font>	
	<br>&nbsp;<b><%=rs("estado") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nº Telefone Fixo ou Celular</font>	
	<br>&nbsp;<b><%=rs("telefone1") %>  </td>
</tr>
</table>	

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">E-mail</font>	
	<br>&nbsp;<b><%=rs("email") %>  </td>
	<td class="fundor" style="border:1px solid #000000;border-top:0px solid #000000" width="130px"><font size="1"><b>Dia de Vencimento</font></td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" width="100px"><font size="1"></font>	
	<br>&nbsp;<b></td>
</tr>
</table>	

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="background-color:black;color:white"><b>&nbsp;TABELA DE VENCIMENTOS:</b></td>
</tr>
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" align="center">
	DATA DE VIGÊNCIA: 1 A 5 = DIA 5 | 6 A 10 = DIA 10 | 11 A 15 = DIA 15 | 16 A 20 = DIA 20 | 21 A 25 = DIA 25 | 26 A 30/31 = DIA 30
	</td>
</tr>
<tr>
	<td class="campor" style="background-color:black;color:white"><b>&nbsp;2 - DADOS DEPENDENTES</b></td>
</tr>
</table>
<%
'********* DEPENDENTES **********
sqld="select m.chapa, m.nrodepend, d.dependente, d.CPF, d.nascimento, d.SEXO, d.CARTAOSUS, d.mae, d.ESTADOCIVIL, d.GRAUPARENTESCO " & _
"from assmed_dep_mudanca m inner join assmed_dep d on d.chapa=m.chapa and d.NRODEPEND=m.nrodepend " & _
"where m.chapa='" & rs("chapa") & "' and '" & dtaccess(rsdatademissao) & "' between ivigencia and fvigencia and m.empresa='I' "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
totaldep=rs2.recordcount
if rs2.recordcount>0 then
do while not rs2.eof
'quadros com info
select case rs2("grauparentesco")
	case "1"
		gp="2"
	case "C"
		gp="1"
	case "5"
		gp="1"
	case else
		gp="3"
end select
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
<td class="campor" style="border-bottom:1px solid #000000;border-left:1px solid #000000" valign="middle" align="center" width="40px">
<%=rs2.absoluteposition%>
</td>
<td class="campor" style="border:0px solid #000000;border-top:0px solid #000000" valign="top">

	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nome Completo:</font>	<br>&nbsp;<b><%=rs2("dependente")%> </td>
	</tr>
	</table>
	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Data de Nascimento:</font>	
		<br>&nbsp;<b><%=day(rs2("nascimento")) & " | " & month(rs2("nascimento")) & " | " & year(rs2("nascimento")) %>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Sexo:</font>	
		<br>&nbsp;<b><%=rs2("sexo") %>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" width="130px"><font size="1">CPF</font>	
		<br>&nbsp;<b><%=rs2("cpf") %>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nº da Declaração de Nascido Vivo</font>	
		<br>&nbsp;<b>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Cartão Nacional de Saúde</font>	
		<br>&nbsp;<b><%=rs2("cartaosus") %>  </td>
	</tr>
	</table>
	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">E C</font>	
		<br>&nbsp;<b><%=rs2("estadocivil")%>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">G P</font>	
		<br>&nbsp;<b><%=gp%>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nome da Mãe Completo</font>	<br>&nbsp;<b><%=rs2("mae")%> </td>
	</tr>
	</table>
</td>
</tr>
</table>

<%
rs2.movenext
loop
end if
rs2.close

ndep=totaldep
if totaldep=0 or totaldep<4 then
'quadros vazios
	for b=1 to 4-totaldep
	ndep=ndep+1
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
<td class="campor" style="border-bottom:1px solid #000000;border-left:1px solid #000000" valign="middle" align="center" width="40px">
<%=ndep%>
</td>
<td class="campor" style="border:0px solid #000000;border-top:0px solid #000000" valign="top">

	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nome Completo:</font>	<br>&nbsp;<b></td>
	</tr>
	</table>
	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Data de Nascimento:</font>	
		<br>&nbsp;<b><%="&nbsp;&nbsp;&nbsp;" & " | " & "&nbsp;&nbsp;&nbsp;" & " | " & "&nbsp;&nbsp;&nbsp;&nbsp;" %>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Sexo:</font>	
		<br>&nbsp;<b>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" width="130px"><font size="1">CPF</font>	
		<br>&nbsp;<b>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nº da Declaração de Nascido Vivo</font>	
		<br>&nbsp;<b>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Cartão Nacional de Saúde</font>	
		<br>&nbsp;<b>  </td>
	</tr>
	</table>
	<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">E C</font>	
		<br>&nbsp;<b>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">G P</font>	
		<br>&nbsp;<b>  </td>
		<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000"><font size="1">Nome da Mãe Completo</font>	<br>&nbsp;<b> </td>
	</tr>
	</table>
</td>
</tr>
</table>
	
<%	
	next
end if
contrib1="01/"&right(contrib1,2)&"/" & left(contrib1,4)
contrib99=right(contrib99,2)&"/" & left(contrib99,4)
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" align="center">
	Legenda: E C: Estado Civil: 1-Solteiro 2-Casado 3-Viuvo 4-Separado 5-Divorciado 6-Outros - GP: Grau de Parentesco: 1-Cônjuge 2-Filho 3-Outros
	</td>
</tr>
<tr>
	<td class="campor" style="background-color:black;color:white"><b>&nbsp;3 - VALOR DA TAXA MENSAL EM R$</b></td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" align="center" width="30%" valign="top"><font size="1">Titular</font>	
	<br>&nbsp;<b>  </td>
	<td class="campor" style="border:0px solid #000000;border-bottom:1px solid #000000" align="center" width="10px"><font size="1">+</font></td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" align="center" width="30%" valign="top"><font size="1">Dependente(s)</font>	
	<br>&nbsp;<b>  </td>
	<td class="campor" style="border:0px solid #000000;border-bottom:1px solid #000000" align="center" width="10px"><font size="1">=</font></td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" align="center" valign="top"><font size="1">Valor - Primeira Mensalidade</font>	
	<br>&nbsp;<b>  </td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="background-color:black;color:white"><b>&nbsp;4 - DADOS SOBRE A CONTRIBUIÇÃO E PERÍODO DE MANUTENÇÃO NO PLANO DE SAÚDE</b></td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" width="150px" style="border:1px solid #000000;border-top:0px solid #000000" rowspan="2" valign="top" align="center"><font size="1">Data do início da<br>Contribuição</font>	
	<br>&nbsp;<b><%=contrib1 %>  </td>
	<td class="campor" width="150px" style="border:1px solid #000000;border-top:0px solid #000000" rowspan="2" valign="top" align="center"><font size="1">Data da comunicação do Aviso<br>Prévio ou Aposentadoria</font>	
	<br>&nbsp;<b><%=rsdatademissao %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" colspan="2" valign="top" align="center"><font size="1">Vigência do Período de Manutenção</font>	</td>
</tr>
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center"><font size="1">Vigência da Cobertura</font>	
	<br>&nbsp;<b><%=contrib1 %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center"><font size="1">Término da Cobertura</font>	
	<br>&nbsp;<b><%=contrib99 %>  </td>
</tr>
</table>	

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" width="305px" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center"><font size="1">Operadoras</font></td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center"><font size="1">Data de Início da Contribuição</font></td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center"><font size="1">Data de Término da Contribuição</font></td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center"><font size="1">Total de Meses</font></td>
</tr>
<%
sqlp="select p.chapa, p.empresa, e.operadora, 'Inicio'=min(p.ivigencia), 'Termino'=max(p.fvigencia), 'Meses'=DATEDIFF(M,min(p.ivigencia),dateadd(d,1,max(p.fvigencia))) " & _
"from assmed_mudanca p inner join assmed_empresa e on e.codigo=p.empresa " & _
"where chapa='" & rs("chapa") & "' and empresa not in ('IP','BP','MP','O','UC','V','T','D') " & _
"group by p.chapa, p.empresa, e.operadora order by MIN(p.ivigencia) "
rs2.Open sqlp, ,adOpenStatic, adLockReadOnly
totalplanos=rs2.recordcount
if rs2.recordcount>0 then
do while not rs2.eof
%>
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="left">&nbsp;<b><%=rs2("operadora") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center">&nbsp;<b><%=rs2("inicio") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center">&nbsp;<b><%=rs2("termino") %>  </td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center">&nbsp;<b><%=rs2("meses") %>  </td>
</tr>
<%
rs2.movenext
loop
end if
rs2.close

if totalplanos=0 or totalplanos<4 then
	for b=1 to 4-totalplanos
%>
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="left">&nbsp;</td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center">&nbsp;</td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center">&nbsp;</td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="center">&nbsp;</td>
</tr>
<%	
	next
end if
%>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="background-color:black;color:white"><b>&nbsp;5 - DECLARAÇÃO</b></td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="left" colspan="2">
	Declaro para os fins do disposto nos artigos 30 e 31 da Lei 9656/98, regulamentados pela RN nº 279 e suas atualizações, ter ciência dos direitos
	e deveres assegurados pela referida Lei e Resoluções e, declaro que OPTEI pela continuidade da condição de beneficiário nas mesmas condições de cobertura
	assistencial existentes durante a vigência do contrato de trabalho, juntamente com os dependentes designados no item 2, assumindo neste ato a responsabilidade
	pelo pagamento integral das mensalidades, observadas as informações constantes no verso deste Termo.
	</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000;" valign="top" align="left" colspan="2">
	</td>
</tr>
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="left" width="33%">
	<font size="1">Local e data</font>	
	<br>&nbsp;<br>&nbsp;
	</td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="left">
	<font size="1">Assinatura do titular ou responsável legal em caso de beneficiário menor de idade</font>	
	<br>&nbsp;<br>&nbsp;
	</td>
</tr>
</table>
</center></div>
<!-- ----------------------------- -->
<DIV style="page-break-after:always"></DIV>

<!-- PAGINA 2 -->

<div align="center">
<center>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campo" height="62px" width="204px" style="border-right:0px"><img src="../images/intermedica_notredame_ans.png" height="60px" width="202px" border="0" alt=""></td>
	<td class="campop" valign="middle" align="center" style="border-left:0px;border-right:0px" ><span style="font-size:14pt"><b><i>Termo de Opção de Continuidade</i></b></span><br>(Conforme Arts. 30 e 31 da Lei 9.656/79 e RN 279)</td>
</tr>
<tr>
	<td colspan="2" class="campor" style="background-color:black;color:white"><b>&nbsp;6 - INFORMAÇÕES GERAIS: OPÇÃO DE CONTINUIDADE NO PLANO DE SAÚDE</b></td>
</tr>
</table>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
<td>
<span style="line-height:150%">
<ul style="list-style-type:disc">
	<li>6.1. Prazos de Direito
	<ul style="list-style-type:none">
	  <li style="margin-left:22px;text-indent:-37px">6.1.1. O prazo para a opção da manutenção da condição de beneficiário é de até 30 (trinta) dias, após o comunicado da empresa, que deverá ser formalizado
	  no ato de comunicação do aviso prévio, a ser cumprido ou indenizado, ou da comunicação da aposentadoria.</li>
	  <li style="margin-left:22px;text-indent:-37px">6.1.2. Demitidos ou exonerados sem justa causa<br>
	  O período de manutenção será igual a 1/3 (um terço) do tempo durante o qual contribuiu para o pagamento da contraprestação pecuniária do plano privado de
	  assistência à saúde, sendo-lhe garantido um período mínimo de 6 (seis) meses e, no máximo de 24 (vinte e quatro) meses.</li>
	  <li style="margin-left:22px;text-indent:-37px">6.1.3. Aposentados:<br>
	  - Com menos de 10 (dez) anos de contribuição: o direito de permanência no plano será proporcional ao tempo de contribuição, à razão de 1 (um) ano para cada
	  ano de contribuição.<br>
	  - Com 10 (dez) anos de contribuição: o direito de permanência no plano será garantido por prazo indeterminado, observadas as condições previstas no item 6.5
	  abaixo.
	  </li>
	</ul>
	<li>6.2. Da Vigência do Período de Manutenção:<br>
	A vigência do período de manutenção será contada a partir da efetivação do cadastro.
	</li>
	<li>6.3. Da Mensalidade
	<ul style="list-style-type:none">
	  <li style="margin-left:22px;text-indent:-37px">6.3.1. O valor da mensalidade corresponderá ao valor integral do plano de assistência à saúde (valor pago pelo empregador
	  + contribuição paga pelo empregado - titular e dependentes), sendo que exclusivamente a 1ª mensalidade total, será acrescida da Taxa de Implantação.</li>
	  <li style="margin-left:22px;text-indent:-37px">6.3.2. As mensalidades deverão ser pagas na data de vencimento, por meio de boleto enviado ao endereço declarado, sob
	  pena de suspensão das coberturas após de 10 (dez) dias do vencimento. O restabelecimento das coberturas se dará em até 72 horas após o pagamento.</li>
	  <li style="margin-left:22px;text-indent:-37px">6.3.3. As mensalidades serão corrigidas na mesma periodicidade e percentuais aplicados a Empresa Contratante.</li>
	 </ul>
	</li>
	<li>6.4. Da Carteira de Identificação:<br>
	A carteira de identificação será enviada ao endereço informado.
	</li>
	<li>6.5. Da Perda do Direito dos Beneficiários titulares e seus dependentes aos direitos dos artigos 30 e 31 da Lei 9656/98:
	<ul style="list-style-type:none">
	  <li style="margin-left:22px;text-indent:-37px">6.5.1. A perda do direito a manutenção da condição de beneficiário, se dará nas seguintes hipóteses:<br>
	  a) pelo término dos prazos estabelecidos na Lei e reproduzidos no item 6.1 deste Termo;<br>
	  b) quando o beneficiário titular for admitido em novo emprego;<br>
	  c) na falta do pagamento da mensalidade do plano em prazo superior a 30 (trinta) dias de seu vencimento;<br>
	  d) por fraude praticada pelo Beneficiário Titular ou dependentes ou devido a inobservância das obrigações estabelecidades na lei ou no Contrato firmado entre
	  Contratante e a Intermédica; ou <br>
	  e) quando do cancelamento do contrato firmado entre a Intermédica e a Contratante.
	  </li>
	 </ul>
	</li>
	<li>6.6. Qualquer alteração nas condições contratuais vigentes com a Contratante, ou na legislação vigente, serão aplicadas no que couber a este Termo,
	mesmo que retroativamente.
	</li>
	<li>6.7. Da documentação obrigatória<br>
	Na opção pela manutenção da condição de beneficiário deverá ser apresentada a documentação informada no Comunicado para Opção de Continuidade, apresentado
	pela empresa contratante.
	</li>
</ul>
</span>
</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-right:1px solid #000000;" valign="top" align="left" colspan="2">
	</td>
</tr>
<tr>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="left" width="33%">
	<font size="1">Local e data</font>	
	<br>&nbsp;<br>&nbsp;
	</td>
	<td class="campor" style="border:1px solid #000000;border-top:0px solid #000000" valign="top" align="left">
	<font size="1">Assinatura do titular ou responsável legal em caso de beneficiário menor de idade</font>	
	<br>&nbsp;<br>&nbsp;
	</td>
</tr>
</table>

</center></div>
<!-- ----------------------------- -->
<DIV style="page-break-after:always"></DIV>

<!-- Pagina 3 -->

<div align="center">
<center>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" style="border:0px solid #000000"><br>
	Nome da Empresa: <u>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO<%for a=1 to 60:response.write "&nbsp;":next%></u>
	</td>
</tr>
<tr>
	<td class="campo" align="center"><br>
	<b>COMUNICADO PARA OPÇÃO PELA CONTINUIDADE ARTs. 30 e 31 DA LEI PLANOS DE SAÚDE - LEI Nº 9656/98</b>
	</td>
</tr>
<tr>
	<td class="campo"><br>
	É garantido aos ex-empregados, demitidos ou exonerados sem justa causa ou aposentados que contribuiram mensalmente para o pagamento da contraprestação
	pecuniária do plano privado de assistência médica em decorrência de vínculo empregatício, o direito de manterem a condição de beneficiários deste plano,
	nas mesmas condições de cobertura de que gozavam quando da vigência do vínculo de emprego, desde que assumam o pagamento integral da respectiva 
	contraprestação pecuniária (valor pago pelo empregador + contribuição paga pelo empregado: titular e dependente(s)).
	<br>
	<br>
	Não são consideradas contribuições:<br>
	a) valores pagos pelo Titular relacionados à contribuição de dependentes e/ou agregados; e<br>
	b) valor pagos pelo Titular correspondentes à coparticipação;<br>
	<br>
	Tal benefício é extensivo aos dependentes inscritos quando da vigência do contrato de trabalho, sendo certo de que serão excluídos do plano ao término
	dos prazos estabelecidos em lei para manutenção do benefício ou na hipótese de perderem a condição de dependência prevista no contrato. O direito de
	opção pela manutenção ou não no plano privado de assistência médica deverá ser manifestado em até 30 (trinta) dias, contados da comunicação do Aviso
	Prévio ou da comunicação da Aposentadoria, por meio do Termo de Opção de Continuidade.<br>
	<br>
	Tendo lido, o <b>COMUNICADO PARA OPÇÃO</b> sobre os direitos dos Arts. 30 e 31 da Lei nº 9.656/98 - Lei de Planos de Saúde, declaro:<br>
	<br>
	(&nbsp;&nbsp;&nbsp;) <b>Opção pela não</b> continuidade da condição de beneficiário no Plano de Assistência à Saúde;<br>
	(&nbsp;&nbsp;&nbsp;) <b>Opção pela continuidade</b> da condição de beneficiário no Plano de Assistência à Saúde;<br>
	<br>
	<b>Caso a opção seja pela continuidade, deverá ser encaminhado a documentação abaixo:</b><br>
	- Cópia deste documento assinado pelo Beneficiário Titular e pelo Responsável Contratante;<br>
	- Cópia do último holerite ou documentos emitidos pela empresa que demonstrem os descontos referentes à contribuição ao Plano de assistência à saúde;<br>
	- Comprovante de residência em nome do titular;<br>
	- Cópias do RG / CPF ou CNH do titular e dependente(s) quando maiores de 18 anos;<br>
	- Cartão SUS (Sistema Único de Saúde) do titular e dependente(s);<br>
	- Cópia do Termo de Rescisão do Contrato de Trabalho;<br>
	- No caso de Aposentado: apresentar além dos documentos citados acima, cópia da Carta de Concessão da Aposentadoria do INSS;<br>
	<br>
	<b>O boleto para pagamento será encaminhado ao beneficiário Titular, após a efetivação de cadastro.</b><br>
	<br>
	<b>Informações Referente ao Desligamento do Funcionário:</b><br>
	<br>
	Nome do Funcionário: <u><%=rs("nome")%><%for a=1 to 75-len(rs("nome")):response.write "&nbsp;":next%></u> &nbsp; CPF: <u>&nbsp;<%=rs("cpf")%>&nbsp;&nbsp;&nbsp;</u><br>
	<br>
	1 - O Beneficiário foi excluído por:<br>
	(&nbsp;&nbsp;&nbsp;) demissão ou exoneração sem justa causa&nbsp; (&nbsp;&nbsp;&nbsp;) aposentadoria&nbsp; (&nbsp;&nbsp;&nbsp;) demissão após aposentado<br>
	<br>
	2 - O beneficiário demitido ou exonerado sem justa causa é um Beneficiário aposentado que continuava trabalhando na Contratante?<br>
	(&nbsp;<%=Sapos%>&nbsp;) Sim&nbsp; (&nbsp;<%=Napos%>&nbsp;) Não<br>
	<br>
	3 - O Beneficiário contribuia para o pagamento do plano privado de assistência à saúde?<br>
	(&nbsp;<%=Scont%>&nbsp;) Sim&nbsp; (&nbsp;<%=Ncont%>&nbsp;) Não<br>
	<br>
	4 - Por quanto tempo o Beneficiário contribuiu para o pagamento do plano privado de assistência à Saúde?<br>
	<u>&nbsp;&nbsp;<%=cano%>&nbsp;&nbsp;</u> anos <u>&nbsp;&nbsp;<%=cmes%>&nbsp;&nbsp;</u> meses<br>
	<br>
	5 - O ex-empregado optou pela sua manutenção como Beneficiário?<br>
	(&nbsp;&nbsp;&nbsp;) Sim&nbsp; (&nbsp;&nbsp;&nbsp;) Não<br>
	<br>
	<b>Importante:</b><br>
	Este documentação deverá ser impresso e encaminhado a Intermédica Sistema de Saúde através do e-mail <b>cadastro.dap@intermedica.com.br</b>, dentro do prazo
	legal para opção, que se refere em até 30 dias da data do COMUNICADO PARA A OPÇÃO ao funcionário.<br>
	<br>
	Data e Local, _______________________________________________<br>
	<br>
	<br>
	<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="670">
	<tr><td class="campo" style="border-top:2px solid #000000" align="center">Assinatura da Contratante sob Carimbo
		</td>
		<td class="campo">&nbsp;&nbsp;</td>
		<td class="campo" style="border-top:2px solid #000000" align="center">Assinatura do Beneficiário Titular
		</td>
	</tr>
	</table>

	</td>
</tr>
</table>

</center></div>

<!-- ----------------------------- -->

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
	<td class=campo>&nbsp;<a href="decl_copart_intermedica.asp?codigo=<%=rs("chapa")%>"><%=rs("nome")%></a></td>
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
'conexao2.close
'set conexao2=nothing
%>
</body>
</html>