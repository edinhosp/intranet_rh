<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a89")="N" or session("a89")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Licen�a sem Remunera��o</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<div align="right">
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690 height=1000>
<!-- cabe�alho -->
<tr>
<td class="campop" height=40 align="center">
	<b>SOLICITA��O DE AFASTAMENTO<br>LICEN�A SEM REMUNERA��O
</td>
</tr>

<!-- corpo do formul�rio -->
<tr>
<td class="campop" valign="top">
<br>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td colspan=2 class="campop" align="left" style="font-size:12pt">�<br>FUNDA��O INSTITUTO DE ENSINO PARA OSASCO<br>&nbsp;</td>
</tr>
<tr>
	<td class="campop" align="left" style="font-size:12pt">[&nbsp;&nbsp;] Pedido Inicial</td>
	<td class="campop" align="right" style="font-size:12pt">[&nbsp;&nbsp;] Prorroga��o</td>
</tr>
</table>
<p style="line-height:35px;font-size:12pt;text-align:justify;margin-top:0px;margin-bottom:0px">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Eu, _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _, funcion�rio desde<br> _ _ _ _/_ _ _ _/_ _ _ _ _, atualmente 
com _ _ _ _ anos de trabalho nesta institui��o, venho solicitar afastamento atrav�s de licen�a sem remunera��o pelo motivo de _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
 _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _.
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;O per�odo da licen�a sem remunera��o inicia-se em _ _ _ _/_ _ _ _/_ _ _ _ _ <span style="font-family:wingdings;font-size:16pt">�</span> e se 
encerra em _ _ _ _/_ _ _ _/_ _ _ _ _ <span style="font-family:wingdings;font-size:16pt">&#130;</span>.</p>

<p style="margin-top:0px;margin-bottom:0px;font-size:8pt;font-weight:bold"><%for a=1 to 24 :response.write "&nbsp;":next%>(deve coincidir com o t�rmino do per�odo letivo)</span>

<p style="line-height:35px;font-size:12pt;text-align:justify;margin-top:0px;margin-bottom:0px">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Estou ciente de que a inten��o de retornar �s atividades dever� ser comunicada � institui��o at� a data de _ _ _ _/_ _ _ _/_ _ _ _ _
<span style="font-family:wingdings;font-size:16pt">&#131;</span>.
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Igualmente, declaro que ao n�o retornar no t�rmino do afastamento ou ao declarar a inten��o de retorno fora do prazo, estarei me
tornando "Demission�rio", conforme o par�grafo quarto da cl�usula 26 (Licen�a sem remunera��o) da Conven��o Coletiva do SINPRO.

<%for a=1 to 1%><br><%next%>
<p style="line-height:20px;font-size:10pt;text-align:justify">
<br><span style="font-family:wingdings;;font-size:16pt">�</span>A data de <u><b>in�cio da licen�a</b></u> deve ser, preferencialmente, <u><b>pr�xima ao fim do per�odo letivo</b></u>.
<br><span style="font-family:wingdings;;font-size:16pt">&#130;</span>A data de <u><b>t�rmino da licen�a</b></u> deve ser sempre <u><b>coincidente com o in�cio do pr�ximo per�odo letivo</b></u>. Exemplo: 31/01/11, 31/07/11, ...
<br><span style="font-family:wingdings;;font-size:16pt">&#131;</span>No m�nimo, 60 dias antes do t�rmino da licen�a.
</td>
</tr>

<!-- cla�sula da licen�a na conven��o -->
<tr>
<td class=campo height=210>
<p style="font-size:8pt;font-family:tahoma;font-weight:normal;background-color:White;font-size-adjust:inherit;font-stretch:inherit;text-align:justify;margin-top:0;margin-bottom:0;text-align:center">CONVEN��O COLETIVA DOS PROFESSORES
<p style="font-size:8pt;font-family:tahoma;font-weight:bold;background-color:Silver;color:Black;margin-top:0;margin-bottom:0">26. LICEN�A SEM REMUNERA��O</p>
<p style="font-size:8pt;font-family:tahoma;font-weight:normal;background-color:White;font-size-adjust:inherit;font-stretch:inherit;text-align:justify;margin-top:0;margin-bottom:0">
O PROFESSOR com mais de 5 (cinco) anos ininterruptos de servi�o na MANTENEDORA ter� direito a licenciar-se, sem direito � remunera��o, por um <u>per�odo m�ximo de 2 (dois) anos</u>, n�o sendo este per�odo de afastamento computado para contagem de tempo de servi�o ou para qualquer outro efeito, inclusive legal. 
<br><b>Par�grafo primeiro</b> � <u>A licen�a ou sua prorroga��o</u> dever� ser comunicada por escrito, � MANTENEDORA, com <u>anteced�ncia m�nima de noventa dias do per�odo letivo</u>, devendo especificar as datas de in�cio e t�rmino do afastamento. A licen�a s� ter� in�cio a partir da data expressa no comunicado, mantendo-se, at� a�, todas as vantagens contratuais. A <u>inten��o de retorno</u> do PROFESSOR � atividade dever� ser comunicada � MANTENEDORA, no m�nimo, <u>sessenta dias antes do t�rmino do afastamento</u>. 
<br><b>Par�grafo segundo</b> � O <u>t�rmino do afastamento</u> dever� coincidir com o in�cio do per�odo letivo. 
<br><b>Par�grafo terceiro</b> � O PROFESSOR que tenha ou exer�a cargo de confian�a dever�, junto com o comunicado de licen�a, solicitar seu desligamento do cargo a partir do in�cio do per�odo de licen�a.
<br><b>Par�grafo quarto</b> � Considera-se <u>demission�rio</u> o PROFESSOR que, ao t�rmino do afastamento, <u>n�o retornar �s atividades docentes</u>. 
<br><b>Par�grafo quinto</b> � Ocorrendo a dispensa sem justa causa ao t�rmino da licen�a, o PROFESSOR n�o ter� direito � Garantia Semestral de Sal�rios, prevista na cl�usula 29 da presente Conven��o. 
</td>
</tr>

<!-- corpo do formul�rio -->
<tr>
<td class=campo height=150>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" height=145>
<tr>
	<td class=campo colspan=5>CI�NCIA E AUTORIZA��O</td>
</tr>
<tr>
	<td class=campo height=30 style="">_ _ _/_ _ _/_ _ _ _</td>
	<td class=campo width="15"></td>
	<td class=campo style="">_ _ _/_ _ _/_ _ _ _</td>
	<td class=campo width="15"></td>
	<td class=campo style="">_ _ _/_ _ _/_ _ _ _</td>
</tr>
<tr>
	<td class=campo width="32%" height=100></td>
	<td class=campo width="15"></td>
	<td class=campo width="32%"></td>
	<td class=campo width="15"></td>
	<td class=campo width="32%"></td>
</tr>
<tr>
	<td class=campo height=25 style="border-top: 1px solid #000000">SOLICITANTE</td>
	<td class=campo width="15"></td>
	<td class=campo style="border-top: 1px solid #000000">COORDENADOR</td>
	<td class=campo width="15"></td>
	<td class=campo style="border-top: 1px solid #000000">REITOR</td>
</tr>
</table>


</td>
</tr>

</table>
</div>
</body>
</html>