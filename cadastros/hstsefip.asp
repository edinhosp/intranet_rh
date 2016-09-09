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
<title>Histórico de Ocorrências SEFIP</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
chapa=request("chapa")

sqla="SELECT CHAPA, DTMUDANCA, CODOCORRENCIA FROM corporerm.dbo.PHSTOCORRENCIA WHERE CHAPA='" & chapa & "' ORDER BY DTMUDANCA"
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=515>
<th class=titulo colspan=5>Histórico de Ocorrências SEFIP</th>
<tr>
	<td class=titulor>Data Mudança</td>
	<td class=titulor>Cód.Ocorrência</td>
	<td class=titulor>Descrição da Ocorrência</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
sql="select descricao from corporerm.dbo.pcodocortrab where codcliente=" & rs("codocorrencia") & ""
'response.write sql
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codocorrencia=rs2("descricao")
rs2.close
%>
<tr>
	<td class="campor"><%=rs("dtmudanca")%></td>
	<td class="campor"><%=rs("codocorrencia")%></td>
	<td class="campor"><%=codocorrencia%></td>
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