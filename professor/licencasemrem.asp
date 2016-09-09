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
<title>Licença sem Remuneração</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<div align="right">
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690 height=1000>
<!-- cabeçalho -->
<tr>
<td class="campop" height=40 align="center">
	<b>SOLICITAÇÃO DE AFASTAMENTO<br>LICENÇA SEM REMUNERAÇÃO
</td>
</tr>

<!-- corpo do formulário -->
<tr>
<td class="campop" valign="top">
<br>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td colspan=2 class="campop" align="left" style="font-size:12pt">À<br>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO<br>&nbsp;</td>
</tr>
<tr>
	<td class="campop" align="left" style="font-size:12pt">[&nbsp;&nbsp;] Pedido Inicial</td>
	<td class="campop" align="right" style="font-size:12pt">[&nbsp;&nbsp;] Prorrogação</td>
</tr>
</table>
<p style="line-height:35px;font-size:12pt;text-align:justify;margin-top:0px;margin-bottom:0px">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Eu, _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _, funcionário desde<br> _ _ _ _/_ _ _ _/_ _ _ _ _, atualmente 
com _ _ _ _ anos de trabalho nesta instituição, venho solicitar afastamento através de licença sem remuneração pelo motivo de _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
 _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _.
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;O período da licença sem remuneração inicia-se em _ _ _ _/_ _ _ _/_ _ _ _ _ <span style="font-family:wingdings;font-size:16pt"></span> e se 
encerra em _ _ _ _/_ _ _ _/_ _ _ _ _ <span style="font-family:wingdings;font-size:16pt">&#130;</span>.</p>

<p style="margin-top:0px;margin-bottom:0px;font-size:8pt;font-weight:bold"><%for a=1 to 24 :response.write "&nbsp;":next%>(deve coincidir com o término do período letivo)</span>

<p style="line-height:35px;font-size:12pt;text-align:justify;margin-top:0px;margin-bottom:0px">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Estou ciente de que a intenção de retornar às atividades deverá ser comunicada à instituição até a data de _ _ _ _/_ _ _ _/_ _ _ _ _
<span style="font-family:wingdings;font-size:16pt">&#131;</span>.
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Igualmente, declaro que ao não retornar no término do afastamento ou ao declarar a intenção de retorno fora do prazo, estarei me
tornando "Demissionário", conforme o parágrafo quarto da cláusula 26 (Licença sem remuneração) da Convenção Coletiva do SINPRO.

<%for a=1 to 1%><br><%next%>
<p style="line-height:20px;font-size:10pt;text-align:justify">
<br><span style="font-family:wingdings;;font-size:16pt"></span>A data de <u><b>início da licença</b></u> deve ser, preferencialmente, <u><b>próxima ao fim do período letivo</b></u>.
<br><span style="font-family:wingdings;;font-size:16pt">&#130;</span>A data de <u><b>término da licença</b></u> deve ser sempre <u><b>coincidente com o início do próximo período letivo</b></u>. Exemplo: 31/01/11, 31/07/11, ...
<br><span style="font-family:wingdings;;font-size:16pt">&#131;</span>No mínimo, 60 dias antes do término da licença.
</td>
</tr>

<!-- claúsula da licença na convenção -->
<tr>
<td class=campo height=210>
<p style="font-size:8pt;font-family:tahoma;font-weight:normal;background-color:White;font-size-adjust:inherit;font-stretch:inherit;text-align:justify;margin-top:0;margin-bottom:0;text-align:center">CONVENÇÃO COLETIVA DOS PROFESSORES
<p style="font-size:8pt;font-family:tahoma;font-weight:bold;background-color:Silver;color:Black;margin-top:0;margin-bottom:0">26. LICENÇA SEM REMUNERAÇÃO</p>
<p style="font-size:8pt;font-family:tahoma;font-weight:normal;background-color:White;font-size-adjust:inherit;font-stretch:inherit;text-align:justify;margin-top:0;margin-bottom:0">
O PROFESSOR com mais de 5 (cinco) anos ininterruptos de serviço na MANTENEDORA terá direito a licenciar-se, sem direito à remuneração, por um <u>período máximo de 2 (dois) anos</u>, não sendo este período de afastamento computado para contagem de tempo de serviço ou para qualquer outro efeito, inclusive legal. 
<br><b>Parágrafo primeiro</b> – <u>A licença ou sua prorrogação</u> deverá ser comunicada por escrito, à MANTENEDORA, com <u>antecedência mínima de noventa dias do período letivo</u>, devendo especificar as datas de início e término do afastamento. A licença só terá início a partir da data expressa no comunicado, mantendo-se, até aí, todas as vantagens contratuais. A <u>intenção de retorno</u> do PROFESSOR à atividade deverá ser comunicada à MANTENEDORA, no mínimo, <u>sessenta dias antes do término do afastamento</u>. 
<br><b>Parágrafo segundo</b> – O <u>término do afastamento</u> deverá coincidir com o início do período letivo. 
<br><b>Parágrafo terceiro</b> – O PROFESSOR que tenha ou exerça cargo de confiança deverá, junto com o comunicado de licença, solicitar seu desligamento do cargo a partir do início do período de licença.
<br><b>Parágrafo quarto</b> – Considera-se <u>demissionário</u> o PROFESSOR que, ao término do afastamento, <u>não retornar às atividades docentes</u>. 
<br><b>Parágrafo quinto</b> – Ocorrendo a dispensa sem justa causa ao término da licença, o PROFESSOR não terá direito à Garantia Semestral de Salários, prevista na cláusula 29 da presente Convenção. 
</td>
</tr>

<!-- corpo do formulário -->
<tr>
<td class=campo height=150>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" height=145>
<tr>
	<td class=campo colspan=5>CIÊNCIA E AUTORIZAÇÃO</td>
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