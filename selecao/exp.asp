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
<title>Vencimento de Contrato de Experiência</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" style="border:1px solid #00000" align="center"><b>AVALIAÇÃO FUNCIONAL DURANTE O PERÍODO DE EXPERIÊNCIA</td>
	<td class="campop"><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width=175></td>
</tr>
<tr><td class=campo colspan=2 height=10  style="border-bottom:1px solid #000000"></td></tr>
<tr>
	<td class=campo colspan=2 align="center">
	<i>Com o objetivo de avaliar se o processo de adaptação e integração do novo colaborador está sendo adequadamente acompanhado <br>
	e se sua capacidade técnica e profissional estão correspondendo às expectativas	desejadas, solicitamos o preenchimento <br>
	deste formulário devolvendo-o impreterivelmente, até a data	estipulada.
	</td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" rowspan=2 align="center" valign="middle" style="border:1px solid"><b>1º PERÍODO</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Vencimento:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Devolver até:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Visto da área de Recursos Humanos</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-left:1px solid">xx/xx/xxxx</td>
	<td class="campop" valign="top" style="border-left:1px solid">xx/xx/xxxx</td>
	<td class="campop" valign="top" style="border-left:1px solid;border-right:1px solid">assinatura</td>
</tr>
<tr>
	<td class="campop" rowspan=2 align="center" valign="middle" style="border:1px solid"><b>2º PERÍODO</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Vencimento:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Devolver até:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Visto da área de Recursos Humanos</td>
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
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Nome do Funcionário</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">RE</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Data de Admissão</td>
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
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Área / Depto</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Superior Hierárquico-Nome</td>
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
	<td class=campo align="center" valign="middle" rowspan=2>ITENS PARA AVALIAÇÃO</td>
	<td class=grupo align="center" valign="middle" colspan=4>1º Período</td>
	<td class=grupo align="center" valign="middle" colspan=4>2º Período</td>
</tr>
<tr>
	<td class="campor" align="center" valign="middle">ÓTIMO</td>
	<td class="campor" align="center" valign="middle">BOM</td>
	<td class="campor" align="center" valign="middle">REGULAR</td>
	<td class="camposs" align="center" valign="middle">ABAIXO DO<br>ESPERADO</td>
	<td class="campor" align="center" valign="middle">ÓTIMO</td>
	<td class="campor" align="center" valign="middle">BOM</td>
	<td class="campor" align="center" valign="middle">REGULAR</td>
	<td class="camposs" align="center" valign="middle">ABAIXO DO<br>ESPERADO</td>
</tr>
<%
dim itens(15)
itens(0)="Relacionamento: Com colega de trabalho"
itens(1)="Relacionamento: Com a Chefia"
itens(2)="Relacionamento: Com clientes internos e externos"
itens(3)="Desempenho na função: Organização no trabalho"
itens(4)="Desempenho na função: Ritmo - Produtividade"
itens(5)="Desempenho na função: Responsável"
itens(6)="Desempenho na função: Conhecimento do seu Trabalho"
itens(7)="Capacidade de Assimilação"
itens(8)="Capacidade de Adaptação"
itens(9)="Postura Profissional / Ética"
itens(10)="Colaboração / Trabalho de Equipe"
itens(11)="Cumprimento do Horário: Pontualidade"
itens(12)="Assiduidade no trabalho: Faltas"
itens(13)="Realização: entrega e qualidade do trabalho"
itens(14)="Apresentação Pessoal"
itens(15)="Cumprimento: observância das normas e procedimentos"

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
	1<br>º<br> <br>P<br>E<br>R<br>Í<br>O<br>D<br>O</td>
	<td class=campo height="25" valign="middle" style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Decisão: </b>
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
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Acompanhamento/Orientação 
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Treinamento em ______________________________________________
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Data da Devolução</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Visto do Superior Hierárquico</td>
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
	2<br>º<br> <br>P<br>E<br>R<br>Í<br>O<br>D<br>O</td>
	<td class=campo height="25" valign="middle" style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Decisão: </b>
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
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Acompanhamento/Orientação 
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;&nbsp;] Treinamento em ______________________________________________
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Data da Devolução</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Visto do Superior Hierárquico</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid">xxxx</td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid">xxxx</td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>



</body>
</html>
