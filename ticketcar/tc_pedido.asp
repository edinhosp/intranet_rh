<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a62")="N" or session("a62")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pedido de Cartão/Senha - TicketCar</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
'dim conexao, conexao2, chapach, rs, rs2
'set conexao=server.createobject ("ADODB.Connection")
'conexao.Open application("conexao")
'sqla="SELECT dc_carga.CURSO FROM dc_carga GROUP BY dc_carga.CURSO;"
'set rs=server.createobject ("ADODB.Recordset")
'Set rs.ActiveConnection = conexao
'rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<td class="campor" style="border-top: 1px solid #000000">
<table border="0" bordercolor=#000000 cellpadding="8" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" align="center" valign="center" height=50 colspan=2><b><font size=3>REQUISIÇÃO PARA TICKET CAR
	</td>
</tr>
<tr>
	<td class="campop" valign=top colspan=2>
	<p style="margin-top: 2; margin-bottom: 2;line-height: 25px">
	Eu, <input type="text" name="nome" size="70" class="subli">, solicito ao UNIFIEO a requerer à TICKET, 
	administradora do cartão TicketCar, a emissão de:
	</td>
</tr>
<tr>
	<td class="campop" valign="center" colspan=2 height=40 style="border: 1px solid #000000">
    (&nbsp;&nbsp;&nbsp;&nbsp;) cartão
	</td>
</tr>
<tr>
	<td class="campop" valign="center" height=40 style="border: 1px solid #000000">
    (&nbsp;&nbsp;&nbsp;&nbsp;) nova senha
	</td>
	<td class="campop" style="border: 1px solid #000000">
    (&nbsp;&nbsp;&nbsp;&nbsp;) reemissão da senha anterior
	</td>
</tr>
<tr>
	<td class="campop" valign=top height=70 colspan=2 style="border: 1px solid #000000">
	Motivo da solicitação:
	</td>
</tr>

<tr>
	<td class="campop" valign=top height=70 colspan=2 style="border: 0px solid #000000">
<br>	Osasco, ______ de _____________________de <%=year(now)%>
	</td>
</tr>


<tr>
	<td class="campop" valign=top height=70 colspan=2 style="border: 0px solid #000000">
<br>
<br>__________________________________________________________
	</td>
</tr>

</table>

<%
'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>
</body>
</html>