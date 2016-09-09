<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Aulas atribuídas e horários disponíveis</title>
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

chapa=request("chapa")
inicio=request("inicio")
	
sqla="SELECT g.*, h.descricao as descricao2 FROM g2ch g, g2defhor h " & _
"WHERE chapa1='" & chapa & "' and '" & dtaccess(inicio) & "' between inicio and termino " & _
"and g.turno=h.codtn and g.diasem=h.codds and g.pos=h.pos and deletada=0 AND H.TIPOCURSO=2 " & _
"order by diasem, turno, g.pos, coddoc " 
sqla="select d.*, p.* from grades_disp d left join ( " & _
"select chapa1, diasem " & _
", 'h0730'=max(case horini when '07:30' then codtur when '07:45' then codtur else null end) " & _
", 'h0820'=max(case horini when '08:20' then codtur when '08:35' then codtur else null end) " & _
", 'h0920'=max(case horini when '09:20' then codtur when '09:35' then codtur else null end) " & _
", 'h1010'=max(case horini when '10:10' then codtur when '10:25' then codtur else null end) " & _
", 'h1110'=max(case horini when '11:10' then codtur when '11:25' then codtur else null end) " & _
", 'h1200'=max(case horini when '12:00' then codtur when '12:15' then codtur else null end) " & _
", 'h1930'=max(case horini when '19:00' then codtur else null end) " & _
", 'h2020'=max(case horini when '19:50' then codtur else null end) " & _
", 'h2120'=max(case horini when '20:50' then codtur else null end) " & _
", 'h2210'=max(case horini when '21:40' then codtur else null end) " & _
"from g2ch where chapa1='" & chapa & "' and '" & dtaccess(inicio) & "' between inicio and termino " & _
"group by chapa1, diasem  " & _
") p on p.chapa1=d.chapa and p.diasem=d.diasem " & _
"where d.chapa='" & chapa & "' "

rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=510>
	<th class=titulo colspan=12>Aulas atribuídas e disponíveis</th>
	<tr>
		<td class=titulor>Dia</td>
		<td class=titulor align="center">07:30</td>
		<td class=titulor align="center">08:20</td>
		<td class=titulor align="center">09:20</td>
		<td class=titulor align="center">10:10</td>
		<td class=titulor align="center">11:10</td>
		<td class=titulor align="center">12:00</td>
		<td class=titulor align="center">-</td>
		<td class=titulor align="center">19:00</td>
		<td class=titulor align="center">19:50</td>
		<td class=titulor align="center">20:50</td>
		<td class=titulor align="center">21:40</td>
	</tr>
<%
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
%>
	<tr>
		<td class="campor"><%=rs.fields(1)%>-<%=weekdayname(rs.fields(1),-1)%></td>
		<%if rs("m01")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h0730")%></td>
		<%if rs("m02")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h0820")%></td>
		<%if rs("m03")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h0920")%></td>
		<%if rs("m04")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h1010")%></td>
		<%if rs("m05")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h1110")%></td>
		<%if rs("m06")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h1200")%></td>
		<td class=fundor align="left" width=5></td>
		<%if rs("n03")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h1930")%></td>
		<%if rs("n04")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h2020")%></td>
		<%if rs("n05")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h2120")%></td>
		<%if rs("n06")=true then estilo="campolr" else estilo="fundor"%><td class=<%=estilo%> align="left"><%=rs("h2210")%></td>

	</tr>
<%
rs.movenext:loop
%>
<%
else
	response.write "<tr><td class=campo colspan=3>Sem aulas atribuídas</td></tr>"
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