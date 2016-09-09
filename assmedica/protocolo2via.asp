<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" --><html>
<%
	'Response.buffer=true
	'Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a81")="N" or session("a81")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Protocolo de entrega</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
%>
<!-- table pagina -->
<table border="0" width="620" height="450">
<!-- table recibo -->
<% for a=1 to 2 %>
<tr><td valign="top" height="440">
<table border="1" bordercolor="#CCCCCC" cellpadding="5" width="600" cellspacing="0">
<tr><td class=titulo colspan="5"><font size="4">Protocolo de Entrega de 2ª Via - Assistência Médica</font></td></tr>
<tr>
	<td colspan="3">Recebi nesta data os seguintes cartões de assistência médica
	da (&nbsp;&nbsp;&nbsp;) Unimed Seguros (&nbsp;&nbsp;&nbsp;) Intermédica.</td>
</tr>
<tr>
	<td class=titulo>Nome</td>
	<td class=titulo>Parentesco</td>
	<td class=titulo>Carteirinha</td>
</tr>
<tr><td width=300 class=campo>&nbsp;</td><td class=campo>&nbsp;</td><td class=campo>&nbsp;</td></tr>
<tr><td width=300  class=campo>&nbsp;</td><td class=campo>&nbsp;</td><td class=campo>&nbsp;</td></tr>
<tr><td width=300  class=campo>&nbsp;</td><td class=campo>&nbsp;</td><td class=campo>&nbsp;</td></tr>
<tr><td width=300  class=campo>&nbsp;</td><td class=campo>&nbsp;</td><td class=campo>&nbsp;</td></tr>
<tr>
	<td colspan="3">
	Autorizo o desconto em folha de pagamento do valor de R$ ________________ por cada 2ª via de 
	cartão de titular e/ou dependente.<br>
	</td>
</tr>
<tr>
	<td colspan="3">
	Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %><br><br><br>
	______________________________________________________<br>
	Nome:<br>Chapa:
	</td>
</tr>

</table>
<!-- table recibo -->
	</td></tr>
<tr><td><p style='margin-top:0; margin-bottom:0'><font size=1>Recursos Humanos - FIEO
<hr></td></tr>
<%
next
%>
</table>
<!-- table pagina -->
<%
%>
</body>
</html>