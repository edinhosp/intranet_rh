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
sql1="select c.chapa, n.nome, n.descricao, n.codsindicato from (select chapa from veiculos where marca like 'moto%' and dttermino is null group by chapa) c, " & _
"(select chapa, nome, codsecao as descricao, codsindicato from grades_novos union all select f.chapa collate database_default, f.nome collate database_default, f.secao collate database_default, f.codsindicato collate database_default from qry_funcionarios f where f.codsituacao<>'D') n " & _
", (select chapa from veiculos_a where getdate()+1 between inicio and termino and (bp=1 or vy=1) group by chapa) e " & _
"where c.chapa=n.chapa and e.chapa=n.chapa order by n.nome "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
linha=0
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=990>
<tr>
	<td class=grupo colspan=6>RELAÇÃO DE ESTACIONAMENTO BRASIL PARK - MOTOS</td>
	<td class=grupo colspan=1 align="right"><%=now()%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
<tr>
	<td class=titulo>#</td>
	<td class=titulo>Código</td>
	<td class=titulo>Nome Func./Professor</td>
	<td class=titulo>Setor</td>
	<td class=titulo>Contr.</td>
	<td class=titulo>Veículos</td>
	<td class=titulo>Dt.Inclusão</td>
</tr>
<%
linha=2
rs.movefirst
do while not rs.eof

linha=linha+1
if rs("codsindicato")="03" then classe="campoa" else classe="campot"
classe="campo"
if linha>37 then
%>
</table>
<DIV style="page-break-after:always"></DIV>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=990>
<tr>
	<td class=grupo colspan=5>RELAÇÃO DE ESTACIONAMENTO BRASIL PARK</td>
	<td class=grupo colspan=1 align="right"><%=now()%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
<tr>
	<td class=titulo>#</td>
	<td class=titulo>Código</td>
	<td class=titulo>Nome Func./Professor</td>
	<td class=titulo>Setor</td>
	<td class=titulo>Contr.</td>
	<td class=titulo>Veículos</td>
	<td class=titulo>Dt.Inclusão</td>
</tr>
<%
linha=2
end if 'linha
%>
<tr>
	<td class=<%=classe%> align="right"><%=rs.absoluteposition%>&nbsp;</td>
	<td class=<%=classe%> ><%=rs("chapa")%></td>
	<td class=<%=classe%> ><%=rs("nome")%></td>
	<td class=<%=classe%> ><%=rs("descricao")%></td>
	<td class=<%=classe%> >
<%
sql2="select cartao from veiculos_a where chapa='" & rs("chapa") & "' and bp=1 and getdate() between inicio and termino "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
do while not rs2.eof
	response.write rs2("cartao") 
rs2.movenext
loop
end if
rs2.close
%>
	</td>
	<td class=<%=classe%> >
<%
dtcadastro=""
sql2="SELECT modelo, placa, dtcadastro FROM veiculos WHERE dttermino Is Null AND chapa='" & rs("chapa") & "' and marca like 'moto%' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
do while not rs2.eof
	response.write rs2("modelo") 
	response.write "&nbsp;-&nbsp;"
	response.write rs2("placa")
	if rs2.recordcount>1 then response.write "&nbsp;/&nbsp;"
	dtcadastro=dtcadastro & rs2("dtcadastro")
	cadastro=rs2("dtcadastro")
rs2.movenext
loop
end if
rs2.close
if year(now())=year(cadastro) then classe2="grupo" else classe2=classe
%>	
	<td class=<%=classe2%> ><%=dtcadastro%></td>
	</td>
</tr>
<%
rs.movenext
loop
rs.close
%>
</table>
<hr>

<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>