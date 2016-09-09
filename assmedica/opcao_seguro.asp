<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a59")="N" or session("a59")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Opção Seguro Metlife</title>
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
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() { form.chapa.value=form.nome.value; }
function chapa1() { form.nome.value=form.chapa.value; }
--></script>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form("B1")="" then 
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para opção do Seguro
<form method="POST" action="opcao_seguro.asp" name="form">
<%
sqla="SELECT f.chapa, f.NOME FROM corporerm.dbo.pfunc AS f " & _
"WHERE f.CODSINDICATO in ('03','01') AND f.CODSITUACAO<>'D' GROUP BY f.chapa, f.NOME ORDER BY f.NOME;"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>">
<select name="nome" class=a onchange="nome1()">
	<option value="00000">===> Ficha em branco <===</option>
<%
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") then temps="selected" else temps=""
%>
	<option value="<%=rs("chapa")%>" <%=temps%> ><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<br>
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<%
end if

if request.form("B1")<>"" then

chapa=request.form("chapa")
sql1="select chapa, nome, admissao, dtnascimento, sexo, estcivil, mae, cpf, cartidentidade, orgemissorident, ufcartident, rua, numero, complemento, " & _
"bairro, cidade, cep, estado, email, telefone1, telefone2, telefone3, secao, funcao, salario, codsindicato " & _
"from qry_funcionarios q where chapa<'10000' AND q.CHAPA='" & chapa & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	session("chapa")=rs("chapa")
	session("chapanome")=rs("nome")
	idade=int((now()-rs("dtnascimento"))/365.25)
else
end if
%>
<%
if request.form("chapa")="00000" or request.form("chapa")="" then
	admissao="" : registro="" : secao="" : funcao="" : nome="" : nascimento=""
	mae="" : estcivil="" : sexo="" : cpf="" : rg="" : orgao="" : rua=""
	numero="" : complemento="" : bairro="" : cidade="" : cep="" : estado=""
	email="" : telefone1="" : telefone2="" : telefone3="" : sf="" : sm="" : salario=1000
	capital=salario*12
	valormensal=capital*0.000482833:valormensal=capital*0.000531
	valormensal2="___________":extenso=""
else
	admissao=rs("admissao") : registro=rs("chapa") : secao=rs("secao")
	funcao=rs("funcao")     : nome=rs("nome")      : nascimento=rs("dtnascimento")
	mae=rs("mae")           : estcivil=rs("estcivil") : sexo=rs("sexo")
	cpf=rs("cpf")           : rg=rs("cartidentidade") : orgao=rs("orgemissorident") & " " & rs("ufcartident")
	rua=rs("rua")           : numero=rs("numero")     : complemento=rs("complemento")
	bairro=rs("bairro")     : cidade=rs("cidade")     : cep=rs("cep")
	estado=rs("estado")     : email=rs("email")       : telefone1=rs("telefone1")
	telefone2=""            : telefone3=rs("telefone2") : salario=rs("salario")
	if sexo="F" then sf="X" else sm="X"
	if rs("codsindicato")="03" then factor=1.225 else factor=1
	salario=cdbl(salario)*factor
	capital=salario*12
	if capital>120000 then capital=120000
	valormensal=capital*0.000482833::valormensal=capital*0.000531
	valormensal=int(valormensal*100)/100
	valormensal2=formatnumber(valormensal,2)
	extenso=" (" & extenso2(Valormensal) & ")"
end if
tamanho=32:tamanho2=30
corborda="#009999"
corborda="#0066cc"
%>

<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=100>
<tr>
	<td valign="center" align="center" valign=middle><font size=3><b>FICHA DE ADESÃO</b></font></td>
	<td valign="center" align="right"  valign=middle><img src="../images/metlife.jpg" border="0"></td>
</tr>
</tr>
<tr><td colspan=2 height=25 class=fundop align="center" style="background='#0066CC';color:white;font-size:14px"><b>COMO ESTÁ SUA PROTEÇÃO HOJE?</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td class=campo>
	<p style="margin-top:0px;margin-bottom:5px;font-size:12px;line-height: 1.5;text-justify: newspaper;text-indent:0px;">
	Para garantir a segurança dos funcionários da FIEO e de seus familiares, a DCG Corretora de Seguros, oferece para todos os professores e
	funcionários administrativos, um plano de seguro de vida que é garantido pela MetLife Brasil, subsidiária da MetLife Inc, a maior seguradora de
	vida dos Estados Unidos com 2,5 trilhões em capitais segurados.
	
	<p style="margin-top:0px;margin-bottom:5px;font-size:12px;line-height: 1.5;text-justify: newspaper;text-indent:0px;">
	Esta parceria tem como como objetivo principal proporcionar aos seus interessados condições especiais para contratação de Seguros de Vida em Grupo
	com coberturas e custos de acordo com sua necessidade financeira.
	
	<p style="margin-top:0px;margin-bottom:5px;font-size:12px;line-height: 1.5;text-justify: newspaper;text-indent:0px;">
	O que é Seguro de Vida em Grupo?

	<p style="margin-top:0px;margin-bottom:5px;font-size:12px;line-height: 1.5;text-justify: newspaper;text-indent:0px;">
	É um plano de seguro coletivo, exclusivo para uma entidade, vantajoso por ser de baixo custo garantindo aos participantes coberturas compatíveis aos
	riscos que os mesmos estão sujeitos.

	<p style="margin-top:0px;margin-bottom:5px;font-size:12px;line-height: 1.5;text-justify: newspaper;text-indent:0px;">
	Principais Vantagens do Seguro

<ul type="disc" style="margin-top:0px;margin-bottom:5px;font-size:12px">
	<li>Sem carência, passando a vigorar a partir da data de adesão;</li>
	<li>Não é preciso de exame médico para adesão ao plano;</li>
	<li>Alto padrão internacional em seguros</li>
</ul>

	<p style="margin-top:0px;margin-bottom:5px;font-size:12px;line-height: 1.5;text-justify: newspaper;text-indent:0px;">
	Coberturas/Assistências

<ul type="disc" style="margin-top:0px;margin-bottom:5px;font-size:12px">
	<li>Morte por qualquer causa, indenização de 12x o salário bruto, limitado a R$ 120.000,00</li>
	<li>Invalidez total ou parcial por acidente, indenização de até 12x o salário bruto, limitado a R$ 120.000,00</li>
	<li>Invalidez total e permanente por doença, indenização de 12x o salário bruto, limitado a R$ 120.000,00, cujo beneficiário é o próprio segurado</li>
	<li>Assistência Individual Funeral, garante aos beneficiários do segurado em caso de falecimento do mesmo, uma assistência para a realização do sepultamento
	no valor máximo de R$ 2.000,00 prestado através da Central de Atendimento pelo telefone 0800-703 54 33.</li>
</ul>

	<p style="margin-top:0px;margin-bottom:5px;font-size:12px;line-height: 1.5;text-justify: newspaper;text-indent:0px;">
	Confira no exemplo, como é fácil proteger-se:

<ul type="disc" style="margin-top:0px;margin-bottom:5px;font-size:12px">
	<li>Salário base de R$ <%=formatnumber(salario,2)%> x 12 = <%=formatnumber(capital,2)%> (Capital Segurado);</li>
	<li>R$ <%=formatnumber(capital,2)%> x 0,0531% = R$ <%=formatnumber(valormensal,2)%> (custo mensal de seu seguro)</li>
</ul>

	</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=40>
<tr>
	<td valign="center" align="left"  valign=middle style="border-bottom:2px dotted #000000">
	<img src="../images/tesoura1.gif" width="56" height="38" border="0" alt="">
	</td>
</tr>
</table>
<br>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=25>
<tr><td colspan=2 height=25 class=fundop align="center" style="background='#0066CC';color:black;font-size:12px"><b>DADOS PESSOAIS</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td class="campor" bordercolor=<%=corborda%> style="border-bottom:0px #000000 solid;">	&nbsp;Nome Completo</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom:1px #000000 solid;">	&nbsp;<%=nome%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td width="50%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:1px #000000 solid">	&nbsp;Data nascimento</td>
	<td width="50%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;">	&nbsp;CPF</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom:1px #000000 solid;border-right:1px #000000 solid">	&nbsp;<%=nascimento%></td>
	<td bordercolor=<%=corborda%> style="border-bottom:1px #000000 solid;">	&nbsp;<%=cpf%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td colspan=2 class="campop" bordercolor=<%=corborda%> style="border-bottom:3px #000000 solid;border-right:0 solid" valign="middle">&nbsp;</td>
</tr>
</table>

<br>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td colspan=2 height=25 class=fundop align="center" style="background='#3366cc';color:black"><b>CONDIÇÕES PARA ADESÃO</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho2%>>
<tr>
	<td class=campo bordercolor=<%=corborda%> style="border-top:0 solid;border-bottom:3px #000000 solid;font-size:12px"><br>
	Estou ciente das condições do seguro e autorizo o desconto mensal de R$ <%=valormensal2%> <%=extenso%>
	em meu pagamento.
	
	<p align="right">Adesão realizada em Osasco, _____ / _____ / _________
	<p align="right">Assinatura do Beneficiário _________________________________
	<br>&nbsp;
	</td>
</tr></table>


<%
rs.close
set rs=nothing

set rs2=nothing
end if ' 

conexao.close
set conexao=nothing
%>
</body>
</html>