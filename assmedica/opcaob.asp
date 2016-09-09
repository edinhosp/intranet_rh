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
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Opção Intermédica</title>
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
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open application("conexao")
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao2

if request.form<>"" then
	impressao=request.form("operadora1")
	temp=0
else
	temp=1
end if

if temp=1 then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Impressão de Opção por assistência médica de funcionário (administrativo ou professor)
<form method="POST" action="opcaob.asp">
<p style="margin-top: 0; margin-bottom: 0">
	<input type="radio" name="operadora1" value="I"> Intermédica<br>
	<input type="radio" name="operadora1" value="U" checked> Unimed Seguros<br>
	<input type="radio" name="operadora1" value="BS" checked> Bradesco Saúde<br>
	<input type="radio" name="operadora1" value="C" checked> Caixa Seguros<br>
	<input type="submit" value="Imprimir" name="B1" class="button"></p>
</form>
<p><b><font color="#FF0000">Atenção para os períodos para mudança de plano:</font></b><p>

<table border="1" bordercolor="#CCCCCC" cellpadding="7" width="400" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td colspan="3" class=grupo>Prazos</td>
</tr>
<tr>
	<td class=campo>Administrativos</td>
	<td class=campo>Unimed Seguros</td>
	<td class=campo>•Na admissão<br>•em Agosto</td>
</tr>
<tr>
    <td class=campo>Professores</td>
    <td class=campo>Intermédica</td>
    <td class=campo>•Na admissão<br>•em Setembro/Outubro</td>
</tr>
</table>

<%
elseif temp=0 then
select case impressao
	case "I"
		dt_inicio="01/10/2003"
		operadora="Intermédica Sistema de Saúde"
		planogratis="EXTRA"
		anterior="SAMCIL"
		sqlp="select valor from assmed_planos where codigo='I' and plano='4-EXTRA'"
		rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
		valor=rs3("valor")
		rs3.close
		tipo="ADMINISTRATIVO/PROFESSOR"
		clausula="cláusula 49 item 5"
		copar=cdbl(valor*0.1)
	case "U"
		dt_inicio="01/08/2010"
		operadora="Unimed Seguros"
		planogratis="BÁSICO"
		anterior="MEDIAL"
		sqlp="select valor from assmed_planos where codigo='U' and plano='BÁSICO'"
		rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
		valor=rs3("valor")
		rs3.close
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(valor*0.1)
	case "BS"
		dt_inicio="01/11/2014"
		operadora="Bradesco Saúde"
		planogratis="Perfil Enfermaria"
		anterior="UNIMED"
		sqlp="select valor from assmed_planos where codigo='BS' and plano='Perfil Enfermaria'"
		rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
		valor=rs3("valor")
		rs3.close
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(valor*0.1)
	case "C"
		dt_inicio="01/02/2016"
		operadora="Caixa Seguros"
		planogratis="Fundamental Enfermaria"
		anterior="BRADESCO"
		sqlp="select valor from assmed_planos where codigo='C' and plano='Fundamental Enfermaria'"
		rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
		valor=rs3("valor")
		rs3.close
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(valor*0.1)
end select
inicial=0
%>
<table border="0" bordercolor="#000000" cellpadding="4" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td colspan=2 class=titulop style="border: 1px solid #000000"><b>OPÇÕES AO PLANO DE ASSISTÊNCIA MÉDICO-HOSPITALAR</b></td>
</tr>
<tr>
	<td class=campo><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class="campop" style="border-bottom: 1px solid #000000"><p style="text-align:justify"><b>1.</b> Não desejo me filiar ao plano de saúde
	proposto, por já estar filiado a plano de saúde em outra instituição ou particular, e para tanto estou renunciando conforme 
	documento escrito à parte.</td>
</tr>
<tr>
	<td class=campo><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class="campop" style="border-bottom: 1px solid #000000"><p style="text-align:justify"><b>2.</b> Desejo optar pelo plano de saúde abaixo 
	assinalado e contribuir na modalidade de co-participação, conforme artigo 30 da Lei nº 9656/98 e <%=clausula%> da Convenção Coletiva 
	de Trabalho, que permite continuar a usufruir do plano de saúde após rescisão do contrato de trabalho sem justa causa, por um 
	período mínimo de 6 meses e máximo de 24 meses, conforme artigo 30 § 1º da referida lei.</td>
</tr>
<tr>
	<td class=campo><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class="campop" style="border-bottom: 1px solid #000000"><p style="text-align:justify"><b>3.</b> Não desejo contribuir para o plano de 
	saúde na modalidade de co-participação, porém desejo optar pelo plano de saúde abaixo assinalado.</td>
</tr>
<tr>
	<td colspan=2 class="campop">	
	&nbsp;<br>_____________________________________<br>
	<%="Chapa:______ - " & string(40,"_") %></p>
	</td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="4" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulop style="border: 1px solid #000000"><b>AUTORIZAÇÃO PARA DESCONTO E/OU INCLUSÃO</b></td>
</tr>
<tr>
	<td class="campop">
	Eu,&nbsp;<%=string(40,"_") %>, desejo por livre e espontânea vontade, optar por
	um plano de assistência médica diferenciado, identificado abaixo:</td>
</tr>
<tr>
	<td class="campop" valign=top align="center">
<%
sqlplano="SELECT codigo, seq, plano, valor, reembolso FROM (select * from assmed_planos where (codigo='I' and seq<=3) or (codigo='U' and seq<=4) or (codigo='BS' and seq<=5) or codigo='C') a " & _
"WHERE codigo='" & impressao & "' ORDER BY seq "
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly

%>	
	<table border="1" bordercolor="#000000" cellpadding="1" width="500" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td align="center">Opção</font></td>
		<td align="center">Planos</font></td>
		<td align="center">Custo</font></td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular (opção 2)</font></td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular (opção 3)</font></td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Dependente</font></td>
	</tr>
<%
rs3.movefirst
do while not rs3.eof
if impressao="I" then desconto3=rs3("valor") else desconto3=rs3("reembolso")
%>
	<tr>
		<td class=campo align="center"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
		<td class="campop">&nbsp;<%=rs3("plano")%></font></td>
		<td class="campop" align="center"><%=formatnumber(rs3("valor"),2)%></td>
		<td class="campop" align="center"><%=formatnumber(rs3("reembolso")+copar,2)%></td>
		<td class="campop" align="center"><%=formatnumber(rs3("reembolso"),2)%></td>
		<td class="campop" align="center"><%=formatnumber(desconto3,2)%></td>
	</tr>
<%
rs3.movenext
loop
%>
	</table>
<%
rs3.close
set rs3=nothing
conexao2.close
set conexao2=nothing
%>
	</td>
</tr>
<tr>
	<td class="campop">Desejo também&nbsp;incluir os meus dependentes legais (esposa, filhos
	até 21 anos) abaixo relacionados:</td>
</tr>
<tr>
	<td class="campop" valign=top align="center">

	<table border="1" bordercolor="#000000" cellpadding="1" width="600" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td class=titulo align="center">Opção</td>
		<td class=titulo width=250 align="center">Nome do Dependente</td>
		<td class=titulo align="center">Grau de&nbsp;<br>Parentesco</td>
		<td class=titulo align="center">Data de&nbsp;<br>Nascimento</td>
		<td class=titulo align="center">Idade</td>
	</tr>
<%for a=1 to 5%>
	<tr>
		<td align="center" rowspan="2"><font size="2">(&nbsp;&nbsp;&nbsp; )</font></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td class="campor" colspan="4">Nome da Mãe do dependente:<font size=2>&nbsp;</td>
	</tr>
<%next%>
	</table>
	</td>
</tr>
<tr>
	<td class="campop"><p align="justify">Autorizo o desconto mensal em meu salário, através da folha de pagamento, da diferença 
	de valores entre o plano de saúde "<%=planogratis%>" a que tenho direito atualmente como <%=tipo%> e o plano acima por mim 
	escolhido através das opções 2 ou 3, mais o valor da modalidade de co-participação, no caso da opção 2. Estou ciente de que 
	a inclusão do(s) meu(s) dependente(s) será paga integralmente por mim, autorizando desde já, o desconto em meu salário. 
	Nesta data a aludida diferença entre os planos mencionados é de R$ __________, devendo sofrer reajuste quando forem 
	corrigidos os valores cobrados da contratante (FIEO) e que segundo critérios estabelecidos pela <%=operadora%>, qualquer 
	alteração no plano só poderei fazer no aniversário do contrato, ou seja, todo mês de <%=monthname(month(dt_inicio))%> de 
	cada ano.</p>
<p>Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %><br>
Autorizo o desconto.</p>
<p>_____________________________________<br>
<%="Chapa:______ - " & string(40,"_") %></p>
	</td>
	</tr>
</table>
<%
end if ' temps
%>
</body>
</html>