<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a95")="N" or session("a95")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Contrato de Experi�ncia</title>
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
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request("codigo")<>"" then
	chapa=request("codigo")
else
	chapa="02379"
	chapa="02474"
end if
dados=0 '0-admissao 1-atual

sql1="select rua, numero, complemento, bairro, cidade, cep, estado from corporerm.dbo.psecao s, corporerm.dbo.pfunc f " & _
"where f.codsecao=s.codigo and f.chapa='" & chapa & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rua=rs("rua"):numero=rs("numero"):complemento=rs("complemento"):bairro=rs("bairro"):cidade=rs("cidade"):cep=rs("cep"):estado=rs("estado")
rs.close
if right(rua,1)="," then endereco=rua & " " & numero else endereco=rua & ", " & numero
if complemento<>"" then endereco=endereco & " - " & complemento
endereco=endereco & " - " & bairro

sql1="select f.nome, p.sexo, p.carteiratrab, p.seriecarttrab, f.dataadmissao, f.codsindicato " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and f.chapa='" & chapa & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly

'*************funcao
if dados=1 then
	sql1="select f.codfuncao, c.nome as funcao from corporerm.dbo.pfunc f, corporerm.dbo.pfuncao c where f.chapa='" & chapa & "' and c.codigo=f.codfuncao"
else
	sql1="select f.codfuncao, c.nome as funcao from corporerm.dbo.pfhstfco f, corporerm.dbo.pfuncao c where c.codigo=f.codfuncao and f.chapa='" & chapa & "' and f.motivo='03' " 'f.dtmudanca=#" & dtaccess(rs("dataadmissao")) & "#"
end if
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
funcao=rs2("funcao")
rs2.close

'*************salario
if rs("codsindicato")<>"03" then
	if dados=1 then
		sql1="select f.salario, f.jornadamensal as jornada, f.codrecebimento from corporerm.dbo.pfunc f where f.chapa='" & chapa & "' "
	else
		sql1="select h.salario, h.jornada, f.codrecebimento from corporerm.dbo.pfhstsal h, corporerm.dbo.pfunc f where f.chapa=h.chapa and f.chapa='" & chapa & "' and h.motivo='10' " 'h.dtmudanca=#" & dtaccess(rs("dataadmissao")) & "# "
	end if
	tiposal=1:forma="m�s"
else 'sindicato=03
	if dados=1 then
		sql1="select f.salario, f.jornadamensal as jornada, f.codrecebimento from corporerm.dbo.pfunc f where f.chapa='" & chapa & "' "
	else
		sql1="select h.codsecao from corporerm.dbo.pfhstsec h where h.motivo='02' and chapa='" & chapa & "' "
		rs2.Open sql1, ,adOpenStatic, adLockReadOnly
		codsecao=rs2("codsecao"):rs2.close
		sql1="select g.sal from g2cursoeve g where g.codccusto='" & codsecao & "' "
		rs2.Open sql1, ,adOpenStatic, adLockReadOnly
		codsal=rs2("sal"):rs2.close
		sql1="select h.salario, h.jornada, f.codrecebimento from corporerm.dbo.pfhstsal h, corporerm.dbo.pfunc f where f.chapa=h.chapa and f.chapa='" & chapa & "' and h.motivo='10' and codevento='" & codsal & "' " 'h.dtmudanca=#" & dtaccess(rs("dataadmissao")) & "# "
	end if
	tiposal=2:forma="hora aula"
end if
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
salario=cdbl(rs2("salario")):jornada=cdbl(rs2("jornada"))/60
rs2.close
if tiposal=2 then salarioc=salario/jornada else salarioc=salario

if rs("sexo")="F" then v1="a" else v1="o"
if rs("sexo")="F" then v2="a" else v2=""
if rs("sexo")="F" then v3="" else v3="o"
%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650" height="950">
<tr><td class=titulop align="center">CONTRATO DE TRABALHO � T�TULO DE EXPERI�NCIA</td></tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	Entre a empresa FIEO FUNDA��O INSTITUTO DE ENSINO PARA OSASCO, com sede em <%=cidade%>/<%=estado%> �
	<%=endereco%>, doravante designada simplesmente EMPREGADORA e <b><%=rs("nome")%></b>, portador<%=v2%> da 
	Carteira de Trabalho e Previd�ncia Social n� <%=rs("carteiratrab")%> s�rie <%=rs("seriecarttrab")%>, a 
	seguir chamad<%=v1%> de apenas EMPREGADO, � celebrado o presente CONTRATO DE EXPERI�NCIA, que ter� vig�ncia �
	partir da data de in�cio de presta��o de servi�o, de acordo com as condi��es a seguir especificadas;</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	1 - Fica o EMPREGADO admitido no quadro de funcion�rios da EMPREGADORA para exercer as fun��es de <b><%=funcao%></b>, 
	mediante a remunera��o de R$ <%=salarioc%> (<%=extenso2(salarioc)%>) por <%=forma%>. A circunst�ncia, por�m, de ser a fun��o
	especificada n�o importa a intransferibilidade do EMPREGADO para outros servi�os, no qual demonstre melhor capacidade de
	adapta��o desde que compat�vel com sua condi��o pessoal.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	2 - O hor�rio de trabalho ser� anotado em sua ficha de registro e a eventual redu��o da jornada, por determina��o da 
	EMPREGADORA, n�o inovar� este ajuste, permanecendo sempre �ntegra a obriga��o do EMPREGADO de cumprir o hor�rio que lhe for 
	determinado, observando o limite legal.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	3 - Obriga-se tamb�m o EMPREGADO a prestar servi�os em horas extraordin�rias, sempre que lhe for determinado pela EMPREGADORA. 
	O EMPREGADO receber� as horas extraordin�rias com o acr�scimo legal, salvo a ocorr�ncia de compensa��o, com a consequente redu��o
	da jornada de trabalho em outro dia.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	4 - Aceita o EMPREGADO, expressamente, a condi��o de prestar servi�os em qualquer dos turnos de trabalho, isto �, tanto durante
	o dia como a noite, desde sem simultaneidade, observadas as prescri��es do assunto, quanto � remunera��o.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	5 - Fica ajustado nos termos do que disp�e o � 1� do art. 469 da Consolida��o das Leis do Trabalho (CLT), que o 
	EMPREGADO acatar� ordem emanada da EMPREGADORA para a presta��o de servi�os tanto na localidade de celebra��o do Contrato de
	Trabalho, como em qualquer outra Cidade, Capital ou Vila do Territ�rio Nacional, quer essa transfer�ncia seja transit�ria, 
	quer seja definitiva.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	6 - Em caso de dano causado pelo EMPREGADO, fica a EMPREGADORA, autorizada a efetivar o desconto da import�ncia correspondente 
	ao preju�zo, no qual far�, com fundamento no � 1� do art. 462 da Consolida��o das Leis do Trabalho (CLT), j� que essa 
	possibilidade fica expressamente prevista em Contrato.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	7 - A justificativa de aus�ncia do EMPREGADO, deve observar a ordem preferencial dos atestados m�dicos estabelecida pelo 
	Decreto n� 27.048 de 12/08/49, art. 12, �� 1� e 2�, que regulamentou a Lei n� 605/49, conforme segue:
	<br>&nbsp;&nbsp;a) m�dico do Instituto Nacional do Seguro Social (INSS);
	<br>&nbsp;&nbsp;b) m�dico da empresa ou por ela designado e pago;
	<br>&nbsp;&nbsp;c) m�dico do Servi�o Social da Ind�stria (SESI) ou do Servi�o Social do Com�rcio (SESC), conforme o caso;
	<br>&nbsp;&nbsp;d) m�dico de reparti��o federal, estadual ou municipal, incumbida de assuntos de higiene ou sa�de;
	<br>&nbsp;&nbsp;e) m�dico do sindicato a que perten�a o EMPREGADO.
	<br>A ordem preferencial estabelecida na Lei n� 605/49 para a justificativa de faltas ao trabalho d� � EMPREGADORA o direito
	de aceitar ou n�o atestados fornecidos por m�dicos particulares.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	8 - O presente contrato, viger� durante 90 (noventa) dias, sendo celebrado para as partes verificarem reciprocamente, a 
	conveni�ncia ou n�o de se vincularem em car�ter definitivo a um Contrato de Trabalho. A EMPREGADORA passando a conhecer as 
	aptid�es do EMPREGADO e suas qualidades pessoais e morais; o EMPREGADO verificando se o ambiente e os m�todos de trabalho 
	atendem � sua conveni�ncia.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	9 - Opera-se a rescis�o do presente Contrato pela decorr�ncia do prazo supra ou por vontade de uma das partes; rescindindo-se 
	por vontade do EMPREGADO ou pela EMPREGADORA com justa causa, nenhuma indeniza��o � devida; rescindindo-se, antes do prazo, 
	pela EMPREGADORA, fica esta obrigada a pagar 50% dos sal�rios devidos at� o final (metade do tempo combinado restante), nos 
	termos do art. 479 da CLT, sem preju�zo do disposto no Reg. do FGTS. Nenhum aviso pr�vio � devido pela rescis�o do presente 
	Contrato.</td>
</tr>
</table>

<DIV style="page-break-after:always"></DIV>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	10 - Na hip�tese deste ajuste transformar-se em Contrato de Prazo Indeterminado, pelo decurso do tempo, continuar�o em plena 
	vig�ncia as cl�usulas de 1 (um) a 7 (sete), enquanto durarem as rela��es do EMPREGADO com a EMPREGADORA.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	E por estarem de pleno acordo, as partes contratantes, assinam o presente Contrato de Experi�ncia em duas vias, ficando a 
	primeira em poder da EMPREGADORA, e a segunda com o EMPREGADO, que dela dar� o competente recibo.</td>
</tr>

<tr><td>&nbsp;</td>
</tr>

<tr><td>
<%
if ct_contrato="" then ct_contrato=formatdatetime(rs("dataadmissao"),2)
dia=day(ct_contrato)
mes=monthname(month(ct_contrato))
ano=year(ct_contrato)
%>
		<p align="left">Osasco,&nbsp;<%=dia & " de " & mes & " de " & ano %></td>
</tr>

<tr><td>&nbsp;</td></tr>

<tr><td>
		<table border="0" width="100%" cellspacing="0">
		<tr><td width="50%">&nbsp;
				<p>_______________________________________<br>Testemunha</td>
			<td width="50%">&nbsp;
				<p>_______________________________________<br>EMPREGADOR</td>
		</tr>
		<tr><td width="50%">&nbsp;
				<p>_______________________________________<br>Testemunha</td>
			<td width="50%">&nbsp;
				<p>_______________________________________<br>EMPREGADO</td>
		</tr>
		</table>
	</td>
</tr>
<%
admissao=rs("dataadmissao")
fim1per=rs("dataadmissao")+44
fim2per=rs("dataadmissao")+89
dia=day(fim1per)
mes=monthname(month(fim1per))
ano=year(fim1per)
%>

<tr><td>&nbsp;</td></tr>
<tr><td class=titulop align="center">TERMO DE PRORROGA��O</td></tr>
<tr><td>&nbsp;</td></tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	Por m�tuo acordo entre as partes, fica o presente contrato de experi�ncia, que deveria vencer nesta data, prorrogado
	at�: <%=formatdatetime(fim2per,2)%>.</td>
</tr>
<tr><td><p align="left">Osasco,&nbsp;<%=dia & " de " & mes & " de " & ano %></td>
</tr>

<tr><td>&nbsp;</td></tr>

<tr><td>
		<table border="0" width="100%" cellspacing="0">
		<tr><td width="50%">&nbsp;
				<p>_______________________________________<br>Testemunha</td>
			<td width="50%">&nbsp;
				<p>_______________________________________<br>EMPREGADOR</td>
		</tr>
		<tr><td width="50%">&nbsp;
				<p>_______________________________________<br>Testemunha</td>
			<td width="50%">&nbsp;
				<p>_______________________________________<br>EMPREGADO</td>
		</tr>
		</table>
	</td>
</tr>


</table>
</div>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<%
%>
</body>
</html>
<%

set rs=nothing
conexao.close
set conexao=nothing
%>