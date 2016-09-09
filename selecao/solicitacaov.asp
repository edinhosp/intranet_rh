<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a50")="N" or session("a50")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Solicitação de Emprego</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<%
espacamento=5
%>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop" width=240 rowspan=2 valign=top align="left"><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td>
	<td class="campop" width=325 height=71><p style="font-family:'Century Gothic';font-size:18pt;margin-top:0;margin-bottom:0"><b>Solicitação de Emprego</b></td>
</tr>
<tr>
	<td class=campo height=20>&nbsp;</td>	
</tr>
</table>

<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo colspan=2 valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b>
	Queira escrever abaixo uma carta de próprio punho, indicando as aptidões que possui e porque julga a vir ser útl a esta 
	Instituição de Ensino Superior.</b></td>
</tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
<tr><td class=campo colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">&nbsp;</td></tr>
</table>

<br>
<br>
<br>
<table border="0" cellpadding="<%=espacamento%>" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class=campo style="border-top:1px solid #000000">Recursos Humanos&nbsp;</td>
	<td class=campo style="border-top:1px solid #000000" align="right">&nbsp;Form. 04/2004-V</td>
</tr>
</table>

<%
%>
</body>
</html>