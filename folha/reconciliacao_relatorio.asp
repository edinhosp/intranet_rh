<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")="N" or session("a48")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Relatório de Pagamentos a Autônomos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">

</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, rt(10), rd(10)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	tipo=request.form("tipo")
	sessao=session.sessionid
	if request.form("anocomp")="0" then
		periodo="r.anocomp>0 AND r.mescomp>0 "
	else
		ano=left(request.form("anocomp"),4)
		mes=right(request.form("anocomp"),2)
		periodo="r.anocomp=" & ano & " AND r.mescomp=" & mes & " "
	end if

sql1="SELECT r.id_rec, r.id_tipo, t.tipo, r.data, r.valor, r.anocomp, r.mescomp, r.obs " & _
"FROM reconciliacao AS r INNER JOIN reconciliacao_eventos AS t ON r.id_tipo = t.id_tipo " & _
"WHERE " & periodo
if tipo<>"0" then sql1=sql1 & " AND r.id_tipo=" & tipo & " "
sql1=sql1 & " ORDER BY t.tipo, r.data "
sql=sql1
'response.write "<br>" & sql
end if

if request.form="" then
%>
<p class=titulo>Geração de relatório de Reconciliação Contábil
<form method="POST" action="reconciliacao_relatorio.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td colspan="2" class=grupo>Parâmetros</td>
</tr>
<tr>
	<td class=titulo>Competência</td>
	<td class=titulo>
<%
sql2="select anocomp, mescomp from reconciliacao group by anocomp, mescomp order by anocomp desc, mescomp desc "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<select size="1" name="anocomp">
	<option value="0">Todas competências</option>
<%
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
valor=numzero(rs("anocomp"),4) & numzero(rs("mescomp"),2)
if rs.absoluteposition=1 then txtsel="selected" else txtsel=""
%>
	<option value="<%=valor%>" <%=txtsel%> ><%=rs("mescomp")&"/"&rs("anocomp")%></option>
<%
rs.movenext:loop
end if
rs.close
%>
</select>
	</td>
</tr>
<tr>
	<td class=titulo>Tipo</td>
	<td class=titulo>
	<select size="1" name="tipo">
	<option value="0">Todos Tipos</option>
<%
sql2="select id_tipo, tipo from reconciliacao_eventos order by tipo"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
%>
	<option value="<%=rs("id_tipo")%>" ><%=rs("tipo")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select>
	</td>
</tr>
</table>
<input type="submit" value="Visualizar relatório" name="Gerar" class="button"></p>
</form>
<p><font color="#FF0000">Configure a página do seu navegador (Internet
Explorer, Netscape, Mozilla, etc) no sentido RETRATO.</font></p>
<%
else
rs.Open sql, ,adOpenStatic, adLockReadOnly
if request.form("anocomp")="0" then
	colspan=5:colspanb=4
else
	colspan=4:colspanb=3
end if
%>
<table border="0" cellpadding="2" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td align="left" ><b>Relação de Pagamentos para Reconciliação Contábil</b></td>
	<td align="center">&nbsp;</td>
	<td align="right"><b><%=rs("mescomp") & "/" & rs("anocomp")%></td>
</tr>
</table>

<table border="0" cellpadding="1" width="690" cellspacing="0" style="border-collapse: collapse">
<tr><td class="campop" colspan=<%=colspan%>>&nbsp;</td></tr>
<tr>
<%
coltipo=200:colobs=330
if request.form("anocomp")="0" then 
response.write "<td class=titulop align=""center"" width=80>Competência</td>"
coltipo=160:colobs=290
end if
%>
	<td class=titulop align="center" width=<%=coltipo%>>Tipo Pagamento</td>
	<td class=titulop align="center" width=<%=colobs%>>Observação</td>
	<td class=titulop align="center" width=70>Data</td>
	<td class=titulop align="center" width=90>Valor</td>
</tr>
<%
linha=2
totaltipo=0:totalgeral=0
rs.movefirst
do while not rs.eof
competencia=rs("mescomp") & "/" & rs("anocomp")

if lasttipo<>rs("tipo") and rs.absoluteposition>1 then
	response.write "<tr>"
	response.write "<td class=titulo style='border-bottom: 1px solid #000000' colspan=" & colspanb & ">Total " & lasttipo &"</td>"
	response.write "<td class=titulo style='border-bottom: 1px solid #000000' align=""right"">" & formatnumber(totaltipo,2) & "</td>"
	response.write "</tr>"
	totaltipo=0:linha=linha+1
end if

if linha>50 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<hr><p style='margin-top:0;margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left' ><b>Relação de Pagamentos para Reconciliação Contábil</b></td>"
	response.write "<td align='center'>&nbsp;</td>"
	response.write "<td align='right'><b>" & rs("mescomp") & "/" & rs("anocomp") & "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border='0' cellpadding='1' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr><td class=""campop"" colspan=4>&nbsp;</td></tr>"
	response.write "<tr>"
	if request.form("anocomp")="0" then response.write "<td class=titulop align=""center"" width=80>Competência</td>"
	response.write "<td class=titulop align=""center"" width=" & coltipo & ">Tipo Pagamento</td>"
	response.write "<td class=titulop align=""center"" width=" & colobs & ">Observação</td>"
	response.write "<td class=titulop align=""center"" width=70>Data</td>"
	response.write "<td class=titulop align=""center"" width=90>Valor </td>"
	response.write "</tr>"
	linha=2
end if
totaltipo =totaltipo +rs("valor")
totalgeral=totalgeral+rs("valor")
%>
<tr>
<%
if request.form("anocomp")="0" then 
response.write "<td class=""campop"" style='border-right:1px solid #000000;border-left:1px solid #000000'>"
response.write competencia
response.write "</td>"
end if
%>
	<td class="campop" style="border-right: 1px solid #000000;border-left: 1px solid #000000"><%=rs("tipo")%></td>
	<td class="campop" style="border-right: 1px solid #000000"><%=rs("obs")%></td>
	<td class="campop" style="border-right: 1px solid #000000" align="center"><%=rs("data")%></td>
	<td class="campop" style="border-right: 1px solid #000000" align="right"><%=formatnumber(rs("valor"),2)%></td>
</tr>
<%
linha=linha+1
lasttipo=rs("tipo")
rs.movenext
loop
rs.close
%>
<tr>
	<td class=titulo style="border-bottom: 1px solid #000000" colspan=<%=colspanb%>>Total <%=lasttipo%></td>
	<td class=titulo style="border-bottom: 1px solid #000000" align="right"><%=formatnumber(totaltipo,2)%></td>
</tr>
<tr>
<%
if request.form("anocomp")="0" then 
response.write "<td class=""campop"" style='border-top: 1px solid #000000'>&nbsp;</td>"
end if
%>
	<td class="campop" style="border-top: 1px solid #000000">&nbsp;</td>
	<td class="campop" style="border-top: 1px solid #000000">&nbsp;</td>
	<td class="campop" style="border-top: 1px solid #000000">&nbsp;</td>
	<td class="campop" style="border-top: 1px solid #000000" align="right"><%=formatnumber(totalgeral,2)%></td>
</tr>
</table>
<%
linha=linha+1
pagina=pagina+1
'response.write "<br>"
response.write "<hr><p style='margin-top:0;margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
'response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 

end if 'request.form
%>
</body>
</html>
<%
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>