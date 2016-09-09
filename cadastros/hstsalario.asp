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
<title>Histórico Salarial</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, sal_anterior(10)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
chapa=request("chapa")

sqla="SELECT * FROM corporerm.dbo.PFHSTSAL WHERE CHAPA='" & chapa & "' ORDER BY DTMUDANCA, nrosalario"
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=515>
<th class=titulo colspan=10>Histórico Salarial</th>
<tr>
	<td class=titulor>Dt.Ref.</td>
	<td class=titulor>Dt.Mudança</td>
	<td class=titulor>Motivo</td>
	<td class=titulor>Desc. Motivo</td>
	<td class=titulor>Nro</td>
	<td class=titulor>Salário</td>
	<td class=titulor>Jornada</td>
	<td class=titulor>Perc.</td>
	<td class=titulor>Hora</td>
	<td class=titulor>Evento</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
'for a=1 to 10:sal_anterior(a)=cdbl(rs("salario")):next
sal_anterior(cint(rs("nrosalario")))=cdbl(rs("salario"))
do while not rs.eof
sql="select descricao from corporerm.dbo.pmotmudsal where codcliente='" & rs("motivo") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then descmotivo=trim(rs2("descricao")) else descmotivo=""
rs2.close
jornadames=cdbl(rs("jornada")/60)
horac=int(rs("jornada")/60)
minutoc=int(((rs("jornada")/60)-horac)*60)
jornadamensal=horac & ":" & numzero(minutoc,2)
if cdbl(rs("salario"))<>0 and jornadames<>0 then hora=formatnumber(cdbl(rs("salario"))/jornadames,2) else hora=0
if sal_anterior(cint(rs("nrosalario")))=0 then sal_anterior(cint(rs("nrosalario")))=cdbl(rs("salario"))
perc=formatpercent((cdbl(rs("salario"))/sal_anterior(cint(rs("nrosalario"))))-1,2)
%>
<tr>
	<td class="campor" align="center"><%=rs("datadereferencia")%></td>
	<td class="campor" align="left"><%=rs("dtmudanca")%></td>
	<td class="campor" align="center"><%=rs("motivo")%></td>
	<td class="campor" align="left"><%=descmotivo%></td>
	<td class="campor" align="center"><%=rs("nrosalario")%></td>
	<td class="campor" align="right"><%=formatnumber(rs("salario"),2)%>&nbsp;</td>
	<td class="campor" align="center"><%=jornadamensal%></td>
	<td class="campor" align="right"><%=perc%></td>
	<td class="campor" align="right"><%=hora%>&nbsp;</td>
	<td class="campor" align="center"><%=rs("codevento")%></td>
</tr>
<%
sal_anterior(cint(rs("nrosalario")))=cdbl(rs("salario"))
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