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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Autoriza��o para Desconto</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width=650>
	<tr><td>
	&nbsp;
<p align="center"><b><font size="+1">AUTORIZA��O PARA DESCONTO</font></b></p>
<p style="margin-top: 0; margin-bottom: 0" align="center">Ref. Assist�ncia M�dica</p>
<p>
<p align=justify style="line-height: 30px;"><font size="2">Eu, <%=string(40,"_")%>, portadora do R.G. n� <%=string(20,"_")%>, desejo incluir meu c�njuge
no plano de sa�de da ________________________, denominado <%=string(20,"_")%>, e autorizo desde j� o desconto mensal em meu 
sal�rio, atrav�s da folha de pagamento, do valor integral do plano.<br>
Estou ciente de que nesta data o aludido valor � de R$ __________, devendo sofrer reajuste quando forem
corrigidos os valores cobrados da contratante (FIEO) e que segundo crit�rios estabelecidos pela Unimed Seguros, qualquer altera��o no plano s� poderei fazer no anivers�rio do contrato, ou seja, todo m�s de
maio de cada ano.
</font></p>
<br>
<div align="center">
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width=600 height=60>
	<tr>
		<td valign=top class=campo>Nome do c�njuge</td>
		<td valign=top class=campo width=80>Data de Nascimento</td>
		<td valign=top class=campo>C.P.F.</td>
	</tr>
</table>
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width=600 height=60>
	<tr>
		<td valign=top class=campo>R.G.</td>
		<td valign=top class=campo>Nome da m�e</td>
	</tr>
</table>

</div>
<p>Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %><br><br>
_____________________________________<br>
<%="Chapa:_______"%></p>

<hr>
<p><font size=1>Observa��o:<br>
-Anexar c�pia de certid�o de casamento civil para inclus�o de marido;<br>
-Anexar c�pia de declara��o em cart�rio de vida em conjunto ou certid�o de filhos em comum para inclus�o
de companheiro.</p>
	</td></tr>
</table>

</body>
</html>