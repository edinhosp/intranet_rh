<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")="N" or session("a87")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Relação de Estacionamento da Brasil Park</title>
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
<p class=titulo>Seleção para impressão da lista de exclusões/inclusões</p>
<form method="POST" action="movimentacao.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=350>
<tr>
	<td class=titulo>Pessoa</td>
</tr>
<tr>
	<td class=titulo>
	<select size="1" name="estac">
		<option value="Todos">Todos</option>
		<option value="bp">Brasil Park</option>
		<option value="vy">Vila Yara (Coral)</option>
		<option value="ns">Narciso</option>
	</select>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=350>
<tr><td align="center" class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3"></td></tr>
</table>
</form>
<hr>

<%
end if 'request.form

if request.form<>"" then
estac=request.form("estac")

sql1="select tipo=case when pabp=1 and bp=0 then 'Exclusão' else case when pabp=0 and bp=1 then 'Inclusão' else '--' end end, " & _
"v.chapa, f.nome, vy, ns, bp, inicio, termino, v.cartao, obs, pavy, pans, pabp, status " & _
"from veiculos_a v, (select chapa, nome from grades_novos union all select chapa collate database_default, nome collate database_default from corporerm.dbo.pfunc) f " & _
"where v.chapa=f.chapa and pabp<>bp and termino='02/28/2017' " & _
"order by case when pabp=1 and bp=0 then 'Exclusão' else case when pabp=0 and bp=1 then 'Inclusão' else '--' end end, v.chapa "

rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=620>
<%
do while not rs.eof
if lasttipo<>rs("tipo") then
	response.write "<tr><td colspan=3 class=""campo"">Relação de Inclusão/Exclusão - BrasilPark</td><td class=campo align=""right"">" & now &"</td></tr>"
	response.write "<tr><td colspan=10 class=""grupo"">" & rs("tipo") & "</td></tr>"
%><tr>
	<td class=titulo># Fieo</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Nº Cartão BP</td>
	<td class=titulo>Obs.</td>
</tr><%
end if
%>
<tr>
	<td class=campo><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("cartao")%></td>
	<td class=campo><%=rs("obs")%></td>
</tr>
<%
lasttipo=rs("tipo")
rs.movenext
loop
%>
</table>
<%
'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'rs.movefirst
'for a=0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
'next
'response.write "</tr>"
'if rs.recordcount>0 then rs.movefirst
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************
%>

<%
end if 'request.form

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>