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
<title>Relação de Estacionamento</title>
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
if request.form<>"" then escolhe=1 else escolhe=0

if escolhe=0 then	
%>
<form name="form" action="lista_placas.asp" method="post">
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=300>
<tr>
	<td class=grupo colspan=5>Seleção para RELAÇÃO DE ESTACIONAMENTO</td>
</tr>
<tr>
	<td class=titulo rowspan=2>Campus</td>
	<td class=fundo><input type="radio" name="campus" value="Todos" checked>Todos</td>
</tr>
<tr>
	<td class=fundo></td>
</tr>
<tr>
	<td class=grupo colspan=5 align="center">
	<input type="submit" value="Visualizar" class="button" name="B1">
	</td>
</tr>
</table>

</form>
<%	
else 'escolhe=1
sql="SELECT v.placa, v.modelo, v.cor, v.chapa as matricula, f.nome, a.vy, a.ns, a.bp, a.jw, f.secao as descricao " & _
"FROM veiculos v, qry_funcionarios f, veiculos_a a " & _
"where v.chapa=f.chapa collate database_default and v.chapa=a.chapa and f.codsituacao<>'D' " & _
"and (v.dttermino is null or v.dttermino='') and (cast(vy as integer)+ns+bp+JW)<>0 and getdate() between a.inicio and a.termino " & _
"ORDER BY placa "
'response.write sql
rs.Open sql, ,adOpenStatic, adLockReadOnly
linha=0
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=670>
<tr>
	<td class=grupo colspan=5>RELAÇÃO DE ESTACIONAMENTO - UNIFIEO</td>
	<td class=grupo colspan=4 align="right" nowrap ><%=now()%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
<tr>
	<td class=titulo rowspan=2>Placa</td>
	<td class=titulo rowspan=2>Veículo</td>
	<td class=titulo rowspan=2>Código</td>
	<td class=titulo rowspan=2>Nome Func./Professor<br>Setor</td>
	<td class=fundo colspan=4 align="center">Estacionamento permitido</td>
</tr>
<tr>
	<td class=fundo align="center">Coral</td>
	<td class=fundo align="center">J.Wilson</td>
	<td class=fundo align="center">Narciso</td>
	<td class=fundo align="center">Br.Park</td>
</tr>
<%
linha=2
rs.movefirst
do while not rs.eof
'if rs("sind")="03" then classe="campoa" else classe="campot"
classe="campor"
usa=cdbl(rs("vy")) + cdbl(rs("ns")) + cdbl(rs("bp")) + cdbl(rs("jw"))
if usa=0 then rs.movenext
if linha>38 then
%>
</table>
<DIV style="page-break-after:always"></DIV>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=670>
<tr>
	<td class=grupo colspan=5>RELAÇÃO DE ESTACIONAMENTO - UNIFIEO</td>
	<td class=grupo colspan=4 align="right" nowrap ><%=now()%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
<tr>
	<td class=titulo rowspan=2>Placa</td>
	<td class=titulo rowspan=2>Veículo</td>
	<td class=titulo rowspan=2>Código</td>
	<td class=titulo rowspan=2>Nome Func./Professor<br>Setor</td>
	<td class=fundo colspan=4 align="center">Estacionamento permitido</td>
</tr>
<tr>
	<td class=fundo align="center">Coral</td>
	<td class=fundo align="center">J.Wilson</td>
	<td class=fundo align="center">Narciso</td>
	<td class=fundo align="center">Br.Park</td>
</tr>
<%
linha=2
end if 'linha
%>
<tr>
	<td class=<%=classe%> nowrap><b><%=rs("placa")%></td>
	<td class=<%=classe%> ><%=rs("modelo")%> / <%=rs("cor")%></td>
	<td class=<%=classe%> ><%=rs("matricula")%></td>
	<td class=<%=classe%> nowrap><b><%=rs("nome")%></b><br>&nbsp;&nbsp;&nbsp;<%=rs("descricao")%></td>
	<td class=<%=classe%> align="center"><%if rs("vy")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center"><%if rs("jw")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center"><%if rs("ns")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=<%=classe%> align="center"><%if rs("bp")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
<%
	linha=linha+1
%>
</tr>
<%
rs.movenext
loop
rs.close
%>
</table>

<%
end if 'escolhe=1


set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>