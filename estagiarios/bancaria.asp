<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")="N" or session("a72")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Relação de Crédito</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form="" then	
%>
<form action="bancaria.asp" method="post" name="form">
<p class=realce>Emissão de Relação de Crédito - Estagiários</p>
<select name="periodo">
<%
sql2="SELECT pf.ANOCOMP, pf.MESCOMP, pf.NROPERIODO " & _
"FROM corporerm.dbo.PFPERFF pf INNER JOIN corporerm.dbo.PFUNC f ON pf.CHAPA=f.CHAPA " & _
"WHERE f.CODTIPO='T' " & _
"GROUP BY pf.ANOCOMP, pf.MESCOMP, pf.NROPERIODO " & _
"ORDER BY pf.ANOCOMP desc, pf.MESCOMP desc, pf.NROPERIODO "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if rs2("mescomp")<10 then mes="0" & rs2("mescomp") else mes=rs2("mescomp")
%>
	<option value="<%=rs2("anocomp") & mes & rs2("nroperiodo")%>"><%=mes & "/" & rs2("anocomp") & " - Periodo: " & rs2("nroperiodo")%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
<br>
<input type="submit" value="Visualizar" class=button>
</form>
<%
else 'request.form
	vartemp=request.form("periodo")
	ano=left(vartemp,4)
	mes=mid(vartemp,5,2)
	per=mid(vartemp,7,2)
	pagina=0:totalger=0
	inicio=1
sql="SELECT pf.CHAPA, pf.ANOCOMP, pf.MESCOMP, pf.NROPERIODO, pf.BASEIRRF, ff.DTPAGTO, Sum([VALOR] * (case provdescbase when 'D' then -1 else 1 end)) AS TOTAL, f.NOME, f.CODSECAO, s.DESCRICAO, f.CONTAPAGAMENTO, f.CODAGENCIAPAGTO " & _
"FROM (((corporerm.dbo.PFPERFF pf INNER JOIN corporerm.dbo.PFUNC f ON pf.CHAPA=f.CHAPA) INNER JOIN corporerm.dbo.PFFINANC ff ON (pf.MESCOMP=ff.MESCOMP) AND (pf.ANOCOMP=ff.ANOCOMP) AND (pf.NROPERIODO=ff.NROPERIODO) AND (pf.CHAPA=ff.CHAPA)) INNER JOIN corporerm.dbo.PSECAO s ON f.CODSECAO=s.CODIGO) INNER JOIN corporerm.dbo.PEVENTO e ON ff.CODEVENTO=e.CODIGO " & _
"where ff.valor>0 and e.provdescbase<>'B' " & _
"GROUP BY pf.CHAPA, pf.ANOCOMP, pf.MESCOMP, pf.NROPERIODO, pf.BASEIRRF, ff.DTPAGTO, f.NOME, f.CODSECAO, s.DESCRICAO, f.CONTAPAGAMENTO, f.CODAGENCIAPAGTO " & _
"HAVING (((pf.ANOCOMP)=" & ano & ") AND ((pf.MESCOMP)=" & mes & ") AND ((pf.NROPERIODO)=" & per & "))  " & _
"ORDER BY f.CODSECAO, f.NOME, pf.ANOCOMP, pf.MESCOMP, pf.NROPERIODO "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellspacing="0" width="650" cellpadding="0" style="border-collapse: collapse">
<tr>
	<td valign=top class="campor" width=45%>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO - FIEO<br>
		AV. FRANZ VOEGELLI, 300<br>
		OSASCO<br>
		CNPJ 73.063.166/0003-92
	</td>
	<td valign=top class="campop" align="center"><b>RELAÇÃO DE CRÉDITO<br>
		<%=MES%>/<%=ANO%>
	</td>
	<td valign=top class="campor" width=15%>Página: <%=pagina+1%><br>
		Data: <%=formatdatetime(now,2)%>&nbsp;&nbsp;<%=formatdatetime(now,4)%><br>
</tr>
</table>
<br>
<table border="0" cellspacing="0" width="650" cellpadding="2" style="border-collapse: collapse">
<tr>
	<td class="campop" align="center" style="border-bottom: 1px solid #000000">Chapa</td>
	<td class="campop" align="center" style="border-bottom: 1px solid #000000">Nome</td>
	<td class="campop" align="center" style="border-bottom: 1px solid #000000">Nº Agência</td>
	<td class="campop" align="center" style="border-bottom: 1px solid #000000">C.Corrente</td>
	<td class="campop" align="center" style="border-bottom: 1px solid #000000">Líquido</td>
<tr>
<%
rs.movefirst
do while not rs.eof
if lastsecao<>rs("codsecao") then
	if inicio=0 then
		response.write "<tr>"
		response.write "<td class=""campop"" colspan=4 align=""right"">Total da Seção</td>"
		response.write "<td class=""campop"" align=""right"" style=""border-top: 1px solid #000000"">" & formatnumber(totalsec,2) & "&nbsp;</td>"
		response.write "</tr>"
		totalsec=0
	end if
	response.write "<tr>"
	response.write "<td class=""campop"" colspan=5 style=""border-top: 1px solid #000000""><i>" & rs("codsecao") & " - " & rs("descricao") & "</td>"
	response.write "</tr>"
end if

%>
<tr>
	<td class="campop"><%=rs("chapa")%></td>
	<td class="campop"><%=rs("nome")%></td>
	<td class="campop"><%=rs("codagenciapagto")%></td>
	<td class="campop"><%=rs("contapagamento")%></td>
	<td class="campop" align="right"><%=formatnumber(rs("total"),2)%>&nbsp;</td>
<tr>
<%
totalsec=totalsec+cdbl(rs("total"))
totalger=totalger+cdbl(rs("total"))
lastsecao=rs("codsecao")
inicio=0
rs.movenext
loop
%>
<tr>
	<td class="campop" colspan=4 align="right">Total da Seção</td>
	<td class="campop" align="right" style="border-top: 1px solid #000000"><%=formatnumber(totalsec,2)%>&nbsp;</td>
</tr>
<tr>
	<td class="campop" colspan=4 align="right"><b>Total Geral</td>
	<td class="campop" align="right" style="border-top: 1px solid #000000"><%=formatnumber(totalger,2)%>&nbsp;</td>
</tr>

</table>

<%
rs.close
end if

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>