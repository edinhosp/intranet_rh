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
<title>Contrato de Experiência</title>
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
	tiposal=1:forma="mês"
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
<tr><td class=titulop align="center">CONTRATO DE TRABALHO À TÍTULO DE EXPERIÊNCIA</td></tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	Entre a empresa FIEO FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO, com sede em <%=cidade%>/<%=estado%> à
	<%=endereco%>, doravante designada simplesmente EMPREGADORA e <b><%=rs("nome")%></b>, portador<%=v2%> da 
	Carteira de Trabalho e Previdência Social nº <%=rs("carteiratrab")%> série <%=rs("seriecarttrab")%>, a 
	seguir chamad<%=v1%> de apenas EMPREGADO, é celebrado o presente CONTRATO DE EXPERIÊNCIA, que terá vigência à
	partir da data de início de prestação de serviço, de acordo com as condições a seguir especificadas;</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	1 - Fica o EMPREGADO admitido no quadro de funcionários da EMPREGADORA para exercer as funções de <b><%=funcao%></b>, 
	mediante a remuneração de R$ <%=salarioc%> (<%=extenso2(salarioc)%>) por <%=forma%>. A circunstância, porém, de ser a função
	especificada não importa a intransferibilidade do EMPREGADO para outros serviços, no qual demonstre melhor capacidade de
	adaptação desde que compatível com sua condição pessoal.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	2 - O horário de trabalho será anotado em sua ficha de registro e a eventual redução da jornada, por determinação da 
	EMPREGADORA, não inovará este ajuste, permanecendo sempre íntegra a obrigação do EMPREGADO de cumprir o horário que lhe for 
	determinado, observando o limite legal.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	3 - Obriga-se também o EMPREGADO a prestar serviços em horas extraordinárias, sempre que lhe for determinado pela EMPREGADORA. 
	O EMPREGADO receberá as horas extraordinárias com o acréscimo legal, salvo a ocorrência de compensação, com a consequente redução
	da jornada de trabalho em outro dia.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	4 - Aceita o EMPREGADO, expressamente, a condição de prestar serviços em qualquer dos turnos de trabalho, isto é, tanto durante
	o dia como a noite, desde sem simultaneidade, observadas as prescrições do assunto, quanto à remuneração.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	5 - Fica ajustado nos termos do que dispõe o § 1º do art. 469 da Consolidação das Leis do Trabalho (CLT), que o 
	EMPREGADO acatará ordem emanada da EMPREGADORA para a prestação de serviços tanto na localidade de celebração do Contrato de
	Trabalho, como em qualquer outra Cidade, Capital ou Vila do Território Nacional, quer essa transferência seja transitória, 
	quer seja definitiva.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	6 - Em caso de dano causado pelo EMPREGADO, fica a EMPREGADORA, autorizada a efetivar o desconto da importância correspondente 
	ao prejuízo, no qual fará, com fundamento no § 1º do art. 462 da Consolidação das Leis do Trabalho (CLT), já que essa 
	possibilidade fica expressamente prevista em Contrato.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	7 - A justificativa de ausência do EMPREGADO, deve observar a ordem preferencial dos atestados médicos estabelecida pelo 
	Decreto nº 27.048 de 12/08/49, art. 12, §§ 1º e 2º, que regulamentou a Lei nº 605/49, conforme segue:
	<br>&nbsp;&nbsp;a) médico do Instituto Nacional do Seguro Social (INSS);
	<br>&nbsp;&nbsp;b) médico da empresa ou por ela designado e pago;
	<br>&nbsp;&nbsp;c) médico do Serviço Social da Indústria (SESI) ou do Serviço Social do Comércio (SESC), conforme o caso;
	<br>&nbsp;&nbsp;d) médico de repartição federal, estadual ou municipal, incumbida de assuntos de higiene ou saúde;
	<br>&nbsp;&nbsp;e) médico do sindicato a que pertença o EMPREGADO.
	<br>A ordem preferencial estabelecida na Lei nº 605/49 para a justificativa de faltas ao trabalho dá à EMPREGADORA o direito
	de aceitar ou não atestados fornecidos por médicos particulares.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	8 - O presente contrato, vigerá durante 90 (noventa) dias, sendo celebrado para as partes verificarem reciprocamente, a 
	conveniência ou não de se vincularem em caráter definitivo a um Contrato de Trabalho. A EMPREGADORA passando a conhecer as 
	aptidões do EMPREGADO e suas qualidades pessoais e morais; o EMPREGADO verificando se o ambiente e os métodos de trabalho 
	atendem à sua conveniência.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	9 - Opera-se a rescisão do presente Contrato pela decorrência do prazo supra ou por vontade de uma das partes; rescindindo-se 
	por vontade do EMPREGADO ou pela EMPREGADORA com justa causa, nenhuma indenização é devida; rescindindo-se, antes do prazo, 
	pela EMPREGADORA, fica esta obrigada a pagar 50% dos salários devidos até o final (metade do tempo combinado restante), nos 
	termos do art. 479 da CLT, sem prejuízo do disposto no Reg. do FGTS. Nenhum aviso prévio é devido pela rescisão do presente 
	Contrato.</td>
</tr>
</table>

<DIV style="page-break-after:always"></DIV>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	10 - Na hipótese deste ajuste transformar-se em Contrato de Prazo Indeterminado, pelo decurso do tempo, continuarão em plena 
	vigência as cláusulas de 1 (um) a 7 (sete), enquanto durarem as relações do EMPREGADO com a EMPREGADORA.</td>
</tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	E por estarem de pleno acordo, as partes contratantes, assinam o presente Contrato de Experiência em duas vias, ficando a 
	primeira em poder da EMPREGADORA, e a segunda com o EMPREGADO, que dela dará o competente recibo.</td>
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
<tr><td class=titulop align="center">TERMO DE PRORROGAÇÃO</td></tr>
<tr><td>&nbsp;</td></tr>

<tr><td class="campop" style="text-align:justify"><p style="line-height:20px">
	Por mútuo acordo entre as partes, fica o presente contrato de experiência, que deveria vencer nesta data, prorrogado
	até: <%=formatdatetime(fim2per,2)%>.</td>
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