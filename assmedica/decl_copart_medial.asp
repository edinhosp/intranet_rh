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
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para Declaração opcional de Plano de Saúde
<form method="POST" action="decl_copart_medial.asp">
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
<tr><td valign="center" align="left" ><img src="../images/medial2.gif" width=75%  border="0"></td>
	<td valign="center" align="right"><font style="font-size:14px"><b>DECLARAÇÃO OPCIONAL DO PLANO DE SAÚDE</b></font></td>
</tr>
<tr>
	<td align="center" valign=top colspan=2 height=20>
	<table border=0><tr><td valign=top>
	<img src="../images/round_square.jpg" border="0">
	</td><td valign=top>
	Exonerado/Demitido - Resolução CONSU nº 20
	</td><td valign=top>
	&nbsp;&nbsp;&nbsp;
	</td><td valign=top>
	<img src="../images/round_square.jpg" border="0">
	</td><td valign=top>
	Aposentado - Resolução CONSU nº 21
	</td></tr></table>
	</td>
</tr>
<tr><td colspan=2 class=titulo align="center" height=15 style="border:1px solid #000000"><b>DADOS CADASTRAIS</td></tr>
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
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border: 1px solid #000000" width="690">
<tr><td valign=top>

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
		<tr><td class="campor">Nome do Ex-empregado</td></tr>
		<tr><td class="campor" style="border-bottom: 1px solid">&nbsp;<%=rs("nome")%></td></tr>
	</table>

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
		<tr><td class="campor" width=18%>Nome da Mãe do Titular</td></tr>
		<tr><td class="campor" style="border-bottom: 1px solid">&nbsp;<%=mae%></td></tr>
	</table>

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
		<tr><td class="campor" colspan=8>Estado Civil</td><td class="campor" width=60% style="border-left:1px solid #000000">Profissão</td></tr>
<%
if rs("estadocivil")="S" then bolaS="X" else bolaS=""
if rs("estadocivil")="C" or rs("estadocivil")="O" then bolaC="X" else bolaC=""
if rs("estadocivil")="V" then bolaV="X" else bolaV=""
if rs("estadocivil")="D" or rs("estadocivil")="I" then bolaD="X" else bolaD=""
%>
		<tr>		
		<td style="border-bottom: 1px solid"><img src="../images/round_square<%=bolas%>.jpg" border="0"></td><td class="campor" style="border-bottom: 1px solid">Solteiro</td>
		<td style="border-bottom: 1px solid"><img src="../images/round_square<%=bolac%>.jpg" border="0"></td><td class="campor" style="border-bottom: 1px solid">Casado</td>		
		<td style="border-bottom: 1px solid"><img src="../images/round_square<%=bolav%>.jpg" border="0"></td><td class="campor" style="border-bottom: 1px solid">Víuvo</td>		
		<td style="border-bottom: 1px solid"><img src="../images/round_square<%=bolad%>.jpg" border="0"></td><td class="campor" style="border-bottom: 1px solid">Divorciado/Separado</td>		
		<td style="border-bottom: 1px solid;border-left:1px solid" class=campo>&nbsp;<%=rs("funcao")%></td>
		</tr>
	</table>
	
	
	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
		<tr><td class="campor">RG</td><td class="campor" style="border-left: 1px solid">CPF</td><td class="campor" colspan=6 style="border-left: 1px solid">Data de Nascto.</td>
			<td class="campor" style="border-left: 1px solid" colspan=2>Idade</td><td class="campor" style="border-left: 1px solid" colspan=4>Sexo</td></tr>
		<tr>
		<td style="border-bottom: 1px solid" class=campo>&nbsp;<%=rs("cartidentidade")%></td>
		<td style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid" class="campor">&nbsp;<%=rs("cpf")%></td>
<%for a=1 to 6%>
		<td align="center" style="border-right: 1px solid;border-bottom: 1px solid" class="campor"><%=mid(dtnasc,a,1)%></td>
<%next%>
		<td style="border-bottom: 1px solid;border-left: 1px solid" class="campor" align="center"><%=mid(numzero(idade,2),1,1)%></td>
		<td style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid" class="campor" align="center"><%=mid(numzero(idade,2),2,1)%></td>
		
<%
if rs("sexo")="F" then bolaF="X" else bolaF=""
if rs("sexo")="M" then bolaM="X" else bolaM=""
%>		
		<td style="border-bottom: 1px solid" align="center"><img src="../images/round_square<%=bolaf%>.jpg" border="0"></td><td style="border-bottom: 1px solid">F</td>
		<td style="border-bottom: 1px solid" align="center"><img src="../images/round_square<%=bolam%>.jpg" border="0"></td><td style="border-bottom: 1px solid">M</td>
		</tr></table>


	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
	<tr>
		<td class="campor" colspan=6 style="border-right: 1px solid">Data de Admissão</td>
		<td class="campor" colspan=2 style="border-right: 1px solid">&nbsp;</td>
		<td class="campor" colspan=6 style="border-right: 1px solid">Data de Desligamento</td>
		<td class="campor" colspan=2 style="border-right: 1px solid">&nbsp;</td>
		<td class="campor" colspan=6>Data da Aposentadoria</td>
	</tr>
<%
if rs("tipodemissao")="2" or rs("tipodemissao")="4" then bolaSJC="X" else bolaSJC=""
if rs("aposentado")="1" and bolaSJC="" then bolaAP="X" else bolaAP=""
%>		
	<tr>
<%for a=1 to 6%>
		<td align="center" style="border-right: 1px solid;border-bottom: 1px solid" class="campor"><%=mid(dtadmissao,a,1)%></td>
<%next%>

		<td style="border-bottom: 1px solid" align="center"><img src="../images/round_square<%=bolaSJC%>.jpg" border="0"></td>
		<td style="border-bottom: 1px solid;border-right: 1px solid" align="left" class="campor">Demissão sem justa causa</td>

<%for a=1 to 6%>
		<td align="center" style="border-right: 1px solid;border-bottom: 1px solid" class="campor"><%=mid(dtdemissao,a,1)%></td>
<%next%>

		<td style="border-bottom: 1px solid" align="center"><img src="../images/round_square<%=bolaAP%>.jpg" border="0"></td>
		<td style="border-bottom: 1px solid;border-right: 1px solid" align="left" class="campor">Aposentadoria</td>

<%for a=1 to 5
caracter=mid(dtaposent,a,1)
if caracter=" " then caracter="&nbsp;&nbsp;" else caracter=mid(dtaposent,a,1)
%>
		<td align="center" style="border-right: 1px solid;border-bottom: 1px solid" class="campor"><%=caracter%></td>
<%next
caracter=mid(dtaposent,6,1)
if caracter=" " then caracter="&nbsp;&nbsp;" else caracter=mid(dtaposent,6,1)
%><td align="center" style="border-bottom: 1px solid" class="campor"><%=caracter%></td>

		</tr></table>

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
	<tr><td class="campor" style="border-right: 1px solid">Endereço Residencial</td>
		<td class="campor" style="border-right: 1px solid">Nº</td>
		<td class="campor">Complemento</td>		
	</tr>
	<tr><td style="border-bottom: 1px solid;border-right: 1px solid" width=60% class="campor">&nbsp;<%=rs("rua")%></td>
		<td style="border-bottom: 1px solid;border-right: 1px solid" class="campor">&nbsp;<%=rs("numero")%></td>
		<td style="border-bottom: 1px solid" class="campor">&nbsp;<%if isnull(rs("complemento")) or rs("complemento")="" then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" else response.write rs("complemento")%></td>
		
		</tr></table>

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
	<tr><td class="campor" style="border-right: 1px solid">Bairro</td>
		<td class="campor" style="border-right: 1px solid">Cidade</td>
		<td class="campor" style="border-right: 1px solid">UF</td>
		<td class="campor" style="border-right: 1px solid">CEP</td>
		<td class="campor">DDD-Telefone</td>
	</tr>
	<tr>
		<td style="border-bottom: 1px solid;border-right: 1px solid" class="campor">&nbsp;<%if isnull(rs("bairro")) or rs("bairro")="" then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" else response.write rs("bairro")%></td>
		<td style="border-bottom: 1px solid;border-right: 1px solid" class="campor">&nbsp;<%=rs("cidade")%></td>
		<td style="border-bottom: 1px solid;border-right: 1px solid" class="campor">&nbsp;<%=rs("estado")%></td>
		<td style="border-bottom: 1px solid;border-right: 1px solid" class="campor">&nbsp;<%=rs("cep")%></td>
		<td style="border-bottom: 1px solid" class=campo>&nbsp;<%=rs("telefone1")%></td>
		</tr></table>
		
	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
	<tr><td class="campor">Nome da Ex-empregadora</td>
	</tr>
	<tr><td style="border-bottom: 1px solid #000000" class="campor" width=80%>&nbsp;FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
		</tr></table>

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
	<tr><td class="campor" colspan=8 style="border-right: 1px solid">Grau de Instrução</td>
		<td class="campor" colspan=4 width=40%>Endereço de Cobrança</td>
	</tr>
	<tr>
<%
gi=rs("grauinstrucao")
if gi="1" or gi="2" or gi="3" or gi="4" or gi="5" then bolai1="X" else bolai1=""
if gi="6" or gi="7" then bolai2="X" else bolai2=""
if gi="8" then bolai3="X" else bolai3=""
if gi="9" or gi="A" or gi="B" or gi="C" or gi="D" or gi="E" or gi="F" or gi="G" or gi="H" then bolai4="X" else bolai4=""
%>		
		<td ><img src="../images/round_square<%=bolai1%>.jpg" border="0"></td><td class="campor">1º Grau</td>
		<td ><img src="../images/round_square<%=bolai2%>.jpg" border="0"></td><td class="campor">2º Grau</td>		
		<td ><img src="../images/round_square<%=bolai3%>.jpg" border="0"></td><td class="campor">Superior Incompleto</td>		
		<td ><img src="../images/round_square<%=bolai4%>.jpg" border="0"></td><td class="campor">Superior Completo</td>		
		<td style="border-left: 1px solid"><img src="../images/round_squareX.jpg" border="0"></td><td class="campor">Residencial</td>		
		<td ><img src="../images/round_square.jpg" border="0"></td><td class="campor">Comercial</td>		
		</tr></table>
</td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="690">
<tr><td class="campop" height=5></td></tr>
<tr><td class=titulo align="center" height=15 style="border:1px solid #000000"><b>DEPENDENTES INSCRITOS QUANDO DA VIGÊNCIA DO CONTRATO DE TRABALHO</td></tr>
</table>
<%
sqld="SELECT d.chapa, d.dependente, d.nascimento, d.sexo, d.parentesco, d.cpf, d.mae, p.empresa, p.plano " & _
"FROM assmed_dep d, assmed_dep_mudanca p WHERE d.id_dep=p.id_dep and p.plano='" & plano & "' " & _
"AND d.chapa='" & rs("chapa") & "' AND '" & dtaccess(rsdatademissao) & "' Between p.ivigencia And p.fvigencia "
rs3.Open sqld, ,adOpenStatic, adLockReadOnly
'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs3.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs3.fields(a).name) & "<br>" & a & "</td>"'
'next
'response.write "</tr>"
'do while not rs3.eof 
'response.write "<tr>"
'for a= 0 to rs3.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs3.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs3.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
dia5=numzero(day(rs3("nascimento")),2)
mes5=numzero(month(rs3("nascimento")),2)
ano5=right(year(rs3("nascimento")),2)
dtnasc=dia5&mes5&ano5
idade=int((now()-rs3("nascimento"))/365.25)
if rs3("cpf")="" or isnull(rs3("cpf")) then cpfd="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" else cpfd=rs3("cpf")
par=rs3("parentesco")
parentesco="04"
if par="Esposa" or par="Esposo" or par="Companheira" or par="Companheiro" then parentesco="02"
if par="Filho" or par="Filha" then parentesco="03"
%>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:0 solid #000000" width="690">
<tr><td valign=middle align="center" style="border-left: 1px solid;border-bottom: 1px solid;border-right: 1px solid;background-color:silver;color:black" width=20><b><%=rs3.absoluteposition%></td>
	<td>
	<table border=0 bordercolor=#000000 width=100% style="border-collapse:collapse">
	<tr><td class="campor" style="border-right: 1px solid">Nome </td>
		<td class="campor" colspan=6 style="border-right: 1px solid" >Data de Nasc.</td>
		<td class="campor" style="border-right: 1px solid">Idade</td>
		<td class="campor" style="border-right: 1px solid">Sexo</td>
		<td class="campor" style="border-right: 1px solid">Parent.(*)</td>
	</tr>
	<tr><td style="border-bottom: 1px solid;border-right: 1px solid" width=44% class="campor">&nbsp;<%=rs3("dependente")%></td>
<%for a=1 to 5%>
		<td align="center" width=2% style="border-right: 1px solid;border-bottom: 1px solid" class="campor"><%=mid(dtnasc,a,1)%></td>
<%next%><td align="center" width=2%  style="border-bottom: 1px solid" class="campor"><%=mid(dtnasc,6,1)%></td>
		<td style="border-bottom: 1px solid;border-left: 1px solid" width=3% class="campor" align="center"><%=idade%></td>
		<td style="border-bottom: 1px solid;border-left: 1px solid" width=2% class="campor" align="center"><%=rs3("sexo")%></td>
		<td style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid" width=5% class="campor" align="center"><%=parentesco%></td>
	</tr></table>	

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
	<tr><td class="campor" style="Border-right: 1px solid">CPF </td>
		<td class="campor" style="border-right: 1px solid">Nome da Mãe do Dependente </td>
	</tr>
	<tr>
		<td style="border-bottom: 1px solid;Border-right: 1px solid" width=20% class="campor">&nbsp;<%=cpfd%></td>
		<td style="border-bottom: 1px solid;border-right: 1px solid" width=55% class="campor">&nbsp;<%=rs3("mae")%></td>
	</tr></table>	
<!-- -->
	</td>
</tr></table>
<%
rs3.movenext
loop
end if 'rs3.recordcount>0

if rs3.recordcount=0 or rs3.recordcount<5 then
for b=rs3.recordcount+1 to 5
%>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:0 solid #000000" width="690">
<tr><td valign=middle align="center" style="border-left: 1px solid;border-bottom: 1px solid;border-right: 1px solid;background-color:silver;color:black" width=20><b><%=b%></td>
	<td>
	<table border=0 bordercolor=#000000 width=100% style="border-collapse:collapse">
	<tr><td class="campor" style="border-right: 1px solid">Nome </td>
		<td class="campor" colspan=6 style="border-right: 1px solid" >Data de Nasc.</td>
		<td class="campor" style="border-right: 1px solid">Idade</td>
		<td class="campor" style="border-right: 1px solid">Sexo</td>
		<td class="campor" style="border-right: 1px solid">Parent.(*)</td>
	</tr>
	<tr><td style="border-bottom: 1px solid;border-right: 1px solid" width=44% class="campor">&nbsp;</td>
<%for a=1 to 5%>
		<td align="center" width=2% style="border-right: 1px solid;border-bottom: 1px solid" class="campor">&nbsp;</td>
<%next%><td align="center" width=2%  style="border-bottom: 1px solid" class="campor">&nbsp;</td>
		<td style="border-bottom: 1px solid;border-left: 1px solid" width=3% class="campor" align="center">&nbsp;</td>
		<td style="border-bottom: 1px solid;border-left: 1px solid" width=2% class="campor" align="center">&nbsp;</td>
		<td style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid" width=5% class="campor" align="center">&nbsp;</td>
	</tr></table>	

	<table border=0 bordercolor="#000000" width=100% style="border-collapse:collapse">
	<tr><td class="campor" style="Border-right: 1px solid">CPF </td>
		<td class="campor" style="border-right: 1px solid">Nome da Mãe do Dependente </td>
	</tr>
	<tr>
		<td style="border-bottom: 1px solid;Border-right: 1px solid" width=20% class="campor">&nbsp;</td>
		<td style="border-bottom: 1px solid;border-right: 1px solid" width=55% class="campor">&nbsp;</td>
	</tr></table>	
<!-- -->
	</td>
</tr></table>
<%
next
end if
rs3.close

sqlp="select count(chapa) as vezes from corporerm.dbo.pffinanc where codevento='052' and chapa='" & rs("chapa") & "' "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
meses=rs3("vezes")
if meses="" or isnull(meses) then meses=0
permanece=meses/3
fica=int(permanece)
if permanece<=6 then fica=6
if permanece>24 then fica=24
rs3.close
dini=dtdemissao
datafinal=dateserial(year(rsdatademissao),month(rsdatademissao)+fica,day(rsdatademissao))
dia6=numzero(day(datafinal),2)
mes6=numzero(month(datafinal),2)
ano6=right(year(datafinal),2)
dfim=dia6&mes6&ano6
sqlv="SELECT b.CHAPA, m.empresa, m.plano, p.valor " & _
"FROM assmed_beneficiario b, assmed_mudanca m, assmed_planos p " & _
"WHERE b.CHAPA=m.chapa and m.empresa=p.codigo AND m.plano=p.plano " & _
"AND b.CHAPA='" & rs("chapa") & "' AND m.empresa Not In ('MP','IP') AND '" & dtaccess(rsdatademissao) & "' Between [ivigencia] And [fvigencia] "
rs3.Open sqlv, ,adOpenStatic, adLockReadOnly
valortitular=rs3("valor")
plano=rs3("plano")
rs3.close
sqld="SELECT d.CHAPA, Sum(p.valor) AS valor " & _
"FROM assmed_dep d, assmed_dep_mudanca m, assmed_planos p " & _
"WHERE d.id_dep=m.id_dep AND m.plano = p.plano AND m.empresa=p.codigo " & _
"AND m.empresa Not In ('MP','IP') AND '" & dtaccess(rsdatademissao) & "' Between [ivigencia] And [fvigencia] " & _
"GROUP BY d.CHAPA HAVING d.CHAPA='" & rs("chapa") & "' "
rs3.Open sqld, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then valordependente=rs3("valor") else valordependente=0
rs3.close
totalpagar=valortitular+valordependente
%>
<table border=0 width="690" style="border-collapse"><tr><td class="campor" align="left">(*)Parentesco:  Cônjuge - 02  Filhos - 03  Outros - 04</td></tr></table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse:collapse" width="690">
<tr><td class=titulo align="center" height=15 style="border:1px solid"><b>TERMO DE RESPONSABILIDADE</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class="campor">
Declaro ter conhecimento que caso opte pelo direito que me é dado, continuarei vinculado ao Contrato vigente da minha ex-empregadora, 
respeitando todas as suas cláusulas e itens, além de respeitar os artigos 30 e 31 da Lei nº 9656/98, alterada pela Medida Provisória 
nº 1801-14 de 17/06/99, bem como as Resoluções CONSU nº 20 e 21 de 23/03/99. Portanto, resolvo:
</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border: 1px solid #000000" width="690">
<tr><td valign=top>
	<table border=0 width=100% style="border-collapse">
		<tr><td colspan=2 class="campor"><b>SOMENTE PARA EXONERADOS/DEMITIDOS</b></td></tr>
		<tr><td><img src="../images/round_square.jpg" border="0"></td><td class="campor">
			Sair imediatamente do plano de assistência à saúde; para tanto estou devolvendo neste instante o cartão de identificação, 
			<b>bem como as de todos os meus dependentes e agregados, se houver.</b>
			</td></tr>
		<tr><td><img src="../images/round_square.jpg" border="0"></td><td class="campor">
			Permanecer no plano de assistência à saúde pelo período de 
			<u style="font-size:11px"><%for c=1 to 5:response.write mid(dini,c,1)&"|":next:response.write mid(dini,6,1)%></u> a 
			<u style="font-size:11px"><%for c=1 to 5:response.write mid(dfim,c,1)&"|":next:response.write mid(dfim,6,1)%></u> respeitando o disposto no § 1º do 
			art. 30 (como <b>exonerado/demitido</b>) da referida Lei, compromentendo-me a efetuar o pagamento da Taxa Mensal de Manutenção 
			no valor de R$ <%=formatnumber(totalpagar,2)%> (<%=extenso2(totalpagar)%>), que é a soma da minha antiga contribuição mais o valor anteriormente de responsabilidade 
			patronal, e inclusive também pagar a co-participação na utilização e/ou franquia, se os mesmos constarem do Contrato da 
			minha ex-empregadora.
			</td></tr>
		</table>
</td></tr></table>

<table><tr><td height=1></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border: 1px solid #000000" width="690">
<tr><td valign=top>
	<table border=0 width=100% style="border-collapse">
		<tr><td colspan=2 class="campor"><b>SOMENTE PARA APOSENTADOS</b></td></tr>
		<tr><td><img src="../images/round_square.jpg" border="0"></td><td class="campor">
			Sair imediatamente do plano de assistência à saúde; para tanto estou devolvendo neste instante o cartão de identificação, 
			<b>bem como as de todos os meus dependentes e agregados, se houver.</b>
			</td></tr>
		<tr><td><img src="../images/round_square.jpg" border="0"></td><td class="campor">
			Permanecer no plano de assistência à saúde pelo período de 
			<u style="font-size:11px"><%for c=1 to 5:response.write mid(dini,c,1)&"|":next:response.write mid(dini,6,1)%></u> a 
			<u style="font-size:11px"><%for c=1 to 5:response.write mid(dfim,c,1)&"|":next:response.write mid(dfim,6,1)%></u> respeitando o disposto no § 1º do 
			art. 31 (como <b>aposentado</b>) da referida Lei, compromentendo-me a efetuar o pagamento da Taxa Mensal de Manutenção 
			no valor de R$ <%=formatnumber(totalpagar,2)%> (<%=extenso2(totalpagar)%>), que é a soma da minha antiga 
			contribuição mais o valor anteriormente de responsabilidade patronal, e inclusive também pagar a co-participação 
			na utilização e/ou franquia, se os mesmos constarem do Contrato da minha ex-empregadora.
			</td></tr>
		<tr><td><img src="../images/round_square.jpg" border="0"></td><td class="campor">
			Permanecer no plano de assistência à saúde como <b>aposentado</b>, por prazo indeterminado, conforme art. 31, caput. da 
			Lei nº 9656/98, sendo de minha responsabilidade a <b>comprovação de 10 (dez) anos de contribuição no plano atual</b>, fornecendo 
			à Medial Saúde todos os documentos necessários que esta solicitar, comprometendo-me ainda a efetuar o pagamento da Taxa 
			Mensal de Manutenção no valor de R$ <%=formatnumber(totalpagar,2)%> (<%=extenso2(totalpagar)%>), que é a soma da 
			minha antiga contribuição mais o valor anteriormente de responsabilidade patronal, e inclusive também pagar a co-participação 
			na utilização e/ou franquia, se os mesmos constarem do Contrato da minha ex-empregadora.
			</td></tr>
		</table>
</td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class="campor">
O(s) cartão(ões) de identificação utilizado(s) por mim e por meu(s) dependente(s) durante o período em que fui empregado, deverá(ão) 
ser por mim devolvido(s), neste ato, para encaminhamento à Medial Saúde que emitirá outros específicos para este plano. Tenho certeza de que meu direito ao art.30 ou 31 da referida Lei, extinguir-se-á nas seguintes condições:
<br>a) minha admissão em novo emprego;
<br>b) o não pagamento da Taxa Mensal de Manutenção (inclusive a cota patronal) dentro dos prazos estabelecidos;
<br>c) o término do período que me é de direito, respeitando o disposto no § 1º dos artigos 30 e 31 da referida Lei; ou
<br>d) o cancelamento do Contrato da minha ex-empregadora com a Medial Saúde. Portanto, estou ciente e concordo, em que os custos 
gerados pela utilização indevida do plano de assistência à saúde, após a perda da condição de usuário, minha e/ou de meu(s) 
dependente(s) e agregado(s), serão de minha inteira responsabilidade. Comprometo-me ainda, a informar imediatamente minha 
ex-empregadora, a minha admissão em um novo emprego, assim como a perda da condição de dependência dos demais membros do grupo 
familiar, caso em que efetuarei a devolução imediata do(s) respectivo(s) cartão(ões) de identificação do Plano de Saúde;
<br>e) Estou ciente de que para obter o direito, todo o meu grupo familiar inscrito quando da vigência do meu contrato de trabalho
não poderá ser alterado, bem como o plano em que estávamos inscritos, ocasião em que apresentarei inclusive, cópia do meu holerith
para a respectiva comprovação do desconto havido pela minha ex-empregadora e, ainda, deverei aderir ao benefício no prazo máximo
de até 30 (trinta) dias contados do meu desligamento. Comprometo-me, finalmente, a atualizar meu endereço sempre que vier ocorrer
qualquer modificação do mesmo;
<br>f) Declaro ainda, ter tomado conhecimento dos Artigos 30 e 31 da Lei 9656/98, assim como, das Resoluções 20 e 21 do CONSU, aqui
anexadas.
</td>
<!-- 
<td width=2% valign=bottom><img src="../images/medial_ans_vertical.gif"></td> 
-->
</tr>
</table>

<table><tr><td height=1></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td height=30 width=200 class="campor" style="border: 1px solid #000000" valign=top>Local e Data<td>
	<td width=5></td>
	<td class="campor" style="border: 1px solid #000000" valign=top>Carimbo e Assinatura da Empresa<td>
	<td width=5></td>
	<td class="campor" style="border: 1px solid #000000" valign=top>Assinatura do Ex-empregado<td>
	<td width=40><img src="../images/abrinq.jpg" width="40" height="30" border="0" alt=""></td>
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