<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a54")="N" or session("a54")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Vencimento de Contrato de Experi�ncia</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" style="border:1px solid #00000" align="center"><b>AVALIA��O FUNCIONAL DURANTE O PER�ODO DE EXPERI�NCIA</td>
	<td class="campop"><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width=175></td>
</tr>
<tr><td class=campo colspan=2 height=10  style="border-bottom:1px solid #000000"></td></tr>
<tr>
	<td class=campo colspan=2 align="center">
	<i>Com o objetivo de avaliar se o processo de adapta��o e integra��o do novo colaborador est� sendo adequadamente acompanhado <br>
	e se sua capacidade t�cnica e profissional est�o correspondendo �s expectativas	desejadas, solicitamos o preenchimento <br>
	deste formul�rio devolvendo-o impreterivelmente, at� a data	estipulada.
	</td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" rowspan=2 align="center" valign="middle" style="border:1px solid"><b>1� PER�ODO</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Vencimento:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Devolver at�:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Visto da �rea de Recursos Humanos</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-left:1px solid">xx/xx/xxxx</td>
	<td class="campop" valign="top" style="border-left:1px solid">xx/xx/xxxx</td>
	<td class="campop" valign="top" style="border-left:1px solid;border-right:1px solid">assinatura</td>
</tr>
<tr>
	<td class="campop" rowspan=2 align="center" valign="middle" style="border:1px solid"><b>2� PER�ODO</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Vencimento:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Devolver at�:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Visto da �rea de Recursos Humanos</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-left:1px solid;border-bottom:1px solid">xx/xx/xxxx</td>
	<td class="campop" valign="top" style="border-left:1px solid;border-bottom:1px solid">xx/xx/xxxx</td>
	<td class="campop" valign="top" style="border-left:1px solid;border-right:1px solid;border-bottom:1px solid">assinatura</td>
</tr>
<tr><td class=campo colspan=4 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Nome do Funcion�rio</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">RE</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Data de Admiss�o</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid">xxxx</td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid">xxxx</td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid">xxxx</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Cargo</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">�rea / Depto</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Superior Hier�rquico-Nome</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid">xxxx</td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid">xxxx</td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid">xxxx</td>
</tr>
<tr><td class=campo colspan=3 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td align="center">

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo align="center" valign="middle" rowspan=2>ITENS PARA AVALIA��O</td>
	<td class=grupo align="center" valign="middle" colspan=4>1� Per�odo</td>
	<td class=grupo align="center" valign="middle" colspan=4>2� Per�odo</td>
</tr>
<tr>
	<td class="campor" align="center" valign="middle">�TIMO</td>
	<td class="campor" align="center" valign="middle">BOM</td>
	<td class="campor" align="center" valign="middle">REGULAR</td>
	<td class="camposs" align="center" valign="middle">ABAIXO DO<br>ESPERADO</td>
	<td class="campor" align="center" valign="middle">�TIMO</td>
	<td class="campor" align="center" valign="middle">BOM</td>
	<td class="campor" align="center" valign="middle">REGULAR</td>
	<td class="camposs" align="center" valign="middle">ABAIXO DO<br>ESPERADO</td>
</tr>
<%
dim itens(15)
itens(0)="Relacionamento: Com colega de trabalho"
itens(1)="Relacionamento: Com a Chefia"
itens(2)="Relacionamento: Com clientes internos e externos"
itens(3)="Desempenho na fun��o: Organiza��o no trabalho"
itens(4)="Desempenho na fun��o: Ritmo - Produtividade"
itens(5)="Desempenho na fun��o: Respons�vel"
itens(6)="Desempenho na fun��o: Conhecimento do seu Trabalho"
itens(7)="Capacidade de Assimila��o"
itens(8)="Capacidade de Adapta��o"
itens(9)="Postura Profissional / �tica"
itens(10)="Colabora��o / Trabalho de Equipe"
itens(11)="Cumprimento do Hor�rio: Pontualidade"
itens(12)="Assiduidade no trabalho: Faltas"
itens(13)="Realiza��o: entrega e qualidade do trabalho"
itens(14)="Apresenta��o Pessoal"
itens(15)="Cumprimento: observ�ncia das normas e procedimentos"

for a=0 to 15
%>
<tr>
	<td class=campo><%=itens(a)%> (<%=len(itens(a))%>)</td>
	<%for b=1 to 4%><td class=campo></td><%next%>
	<%for b=1 to 4%><td class=campo></td><%next%>
</tr>
<%
next
%>
</table>
</td></tr></table>

<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=grupo align="center" valign="middle" rowspan=6>
	1<br>�<br> <br>P<br>E<br>R<br>�<br>O<br>D<br>O</td>
	<td class=campo height="25" valign="middle" style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Decis�o: </b>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Prorrogar 
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Dispensar
	</td>
</tr>
<tr><td class="campor" height="25" valign="top" style="border-bottom:1px dotted;border-right:1px solid">Justificar</td></tr>
<tr><td class=campo style="border-bottom:1px solid;border-right:1px solid">&nbsp;</td></tr>
<tr><td class="campor" height="25" valign="top" style="border-bottom:1px dotted;border-right:1px solid">Pontos a serem melhorados ou considerados</td></tr>
<tr><td class=campo style="border-bottom:1px solid;border-right:1px solid">&nbsp;</td></tr>
<tr><td class=campo height="25" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	&nbsp;&nbsp;&nbsp;<b>Por meio de </b>
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Acompanhamento/Orienta��o 
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Treinamento em ______________________________________________
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Data da Devolu��o</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Visto do Superior Hier�rquico</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid">xxxx</td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid">xxxx</td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=grupo align="center" valign="middle" rowspan=6>
	2<br>�<br> <br>P<br>E<br>R<br>�<br>O<br>D<br>O</td>
	<td class=campo height="25" valign="middle" style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Decis�o: </b>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Prorrogar 
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Dispensar
	</td>
</tr>
<tr><td class="campor" height="25" valign="top" style="border-bottom:1px dotted;border-right:1px solid">Justificar</td></tr>
<tr><td class=campo style="border-bottom:1px solid;border-right:1px solid">&nbsp;</td></tr>
<tr><td class="campor" height="25" valign="top" style="border-bottom:1px dotted;border-right:1px solid">Pontos a serem melhorados ou considerados</td></tr>
<tr><td class=campo style="border-bottom:1px solid;border-right:1px solid">&nbsp;</td></tr>
<tr><td class=campo height="25" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	&nbsp;&nbsp;&nbsp;<b>Por meio de </b>
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Acompanhamento/Orienta��o 
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Treinamento em ______________________________________________
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Data da Devolu��o</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Visto do Superior Hier�rquico</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid">xxxx</td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid">xxxx</td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>



</body>
</html>
