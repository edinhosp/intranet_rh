<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a35")="N" or session("a35")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Histórico de Empréstimos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
chapa=request("chapa")
codigo=request("codigo")

sqla="select h.CHAPA, h.CODIGO, h.ANOCOMP, h.MESCOMP, h.NROPERIODO, h.TIPO, h.VALOR " & _
"from corporerm.dbo.PFDESCEMPRT h where h.CHAPA='" & chapa & "' and h.CODIGO='" & codigo & "' order by h.ANOCOMP, h.MESCOMP, h.NROPERIODO "
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=515>
<th class=titulo colspan=6>Histórico de Empréstimos</th>
<tr>
	<td class=titulor>Código</td>
	<td class=titulor>Ano</td>
	<td class=titulor>Mês</td>
	<td class=titulor>Período</td>
	<td class=titulor>Tipo</td>
	<td class=titulor>Valor</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class="campor"><%=rs("codigo")%></td>
	<td class="campor"><%=rs("anocomp")%></td>
	<td class="campor"><%=rs("mescomp")%></td>
	<td class="campor"><%=rs("nroperiodo")%></td>
	<td class="campor"><%=rs("tipo")%></td>
	<td class="campor"><%=rs("valor")%></td>
</tr>
<%
rs.movenext
loop
else
	response.write "<tr><td class=campo colspan=3>Sem lançamentos cadastrados</td></tr>"
end if
%>
</table>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>