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
<title>Histórico de Provisões</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
chapa=request("chapa")

sqla="SELECT * from corporerm.dbo.pfhstprov WHERE CHAPA='" & chapa & "' ORDER BY ano, mes"
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
'set rs2=server.createobject ("ADODB.Recordset")
'Set rs2.ActiveConnection = conexao
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=515>
<th class=titulo colspan=8>Histórico de Provisões</th>
<tr>
	<td class=titulor align="center">Ano</td>
	<td class=titulor align="center">Mês</td>
	<td class=titulor align="center">Avos 13</td>
	<td class=titulor align="center">Vr.Prov.13º</td>
	<td class=titulor align="center">Avos prop</td>
	<td class=titulor align="center">Avos venc</td>
	<td class=titulor align="center">Venc.Férias</td>
	<td class=titulor align="center">Vr.Prov.Férias</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
'sql="select descricao from pcodocortrab where codcliente=" & rs("codocorrencia") & ""
'response.write sql
'rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codocorrencia=rs2("descricao")
'rs2.close
valor13=rs("valprov13"):if valor13="" or isnull(valor13) then valor13=0
valorfer=rs("valprovfer"):if valorfer="" or isnull(valorfer) then valorfer=0
%>
<tr>
	<td class="campor" align="center"><%=rs("ano")%></td>
	<td class="campor" align="center"><%=rs("mes")%></td>
	<td class="campor" align="center"><%=rs("nroavos13dec")%></td>
	<td class="campor" align="right"><%=formatnumber(valor13,2)%>&nbsp;&nbsp;</td>
	<td class="campor" align="center"><%=rs("nroavosproporcdec")%></td>
	<td class="campor" align="center"><%=rs("nroavosvencferdec")%></td>
	<td class="campor" align="center"><%=rs("dtvencfer")%></td>
	<td class="campor" align="right"><%=formatnumber(valorfer,2)%>&nbsp;&nbsp;</td>
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