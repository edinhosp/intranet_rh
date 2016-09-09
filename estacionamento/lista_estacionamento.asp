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
if request.form<>"" then escolhe=1 else escolhe=0

if escolhe=0 then	
%>
<form name="form" action="lista_estacionamento.asp" method="post">
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=300>
<tr>
	<td class=grupo colspan=5>Seleção para RELAÇÃO DE ESTACIONAMENTO</td>
</tr>
<tr>
	<td class=titulo rowspan=3>Campus</td>
	<td class=fundo><input type="radio" name="campus" value="VY" checked> Vila Yara</td>
</tr>
<tr>
	<td class=fundo><input type="radio" name="campus" value="NS"> Narciso</td>
</tr>
<tr>
	<td class=fundo><input type="radio" name="campus" value="JW"> Jd.Wilson</td>
</tr>
<tr>
	<td class=titulo rowspan=2>Ordem</td>
	<td class=fundo><input type="radio" name="ordem" value="nome" checked> por Nome</td>
</tr>
<tr>
	<td class=fundo><input type="radio" name="ordem" value="chapa"> por Código</td>
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
if request.form("campus")="VY" then sql0="a.vy "
if request.form("campus")="NS" then sql0="a.ns "
if request.form("campus")="JW" then sql0="a.jw "

sql1="SELECT v.chapa, f.CODSITUACAO, " & sql0 & ", Count(v.chapa) AS qt, f.NOME, f.DIASUTPROXMES, f.CODSECAO, f.secao as 'DESCRICAO' " & _
"FROM ((veiculos v INNER JOIN qry_funcionarios f ON v.chapa = f.CHAPA collate database_default) ) INNER JOIN veiculos_a AS a ON f.chapa collate database_default= a.chapa " & _
"WHERE (v.dttermino Is Null or v.dttermino='') and getdate() between a.inicio and a.termino " & _
"GROUP BY v.chapa, f.CODSITUACAO, " & sql0 & ", f.NOME, f.DIASUTPROXMES, f.CODSECAO, f.secao " & _
"HAVING f.CODSITUACAO<>'D' AND " & sql0 & "=1 " 'AND (f.DIASUTPROXMES=0 Or f.DIASUTPROXMES Is Null) "

'if request.form("campus")="VY" then sql1a=" AND a.vy=1 "
'if request.form("campus")="NS" then sql1a=" AND a.ns=1 "
'if request.form("campus")="JW" then sql1a=" AND a.jw=1 "
if request.form("ordem")="nome"  then sql1b=" ORDER BY f.nome "
if request.form("ordem")="chapa" then sql1b=" ORDER BY v.chapa "
	sql=sql1 & sql1a & sql1b
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
linha=0
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650>
<tr>
	<td class=grupo colspan=4>RELAÇÃO DE ESTACIONAMENTO - CAMPUS <%=ucase(request.form("campus"))%></td>
	<td class=grupo colspan=2 align="right" nowrap ><%=now()%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
<tr>
	<td class=titulo>#</td>
	<td class=titulo>Código</td>
	<td class=titulo>Nome Func./Professor</td>
	<td class=titulo>Placa</td>
	<td class=titulo>Veículo</td>
	<td class=titulo>Cor</td>
</tr>
<%
linha=2
rs.movefirst
do while not rs.eof
'if rs("sind")="03" then classe="campoa" else classe="campot"
classe="campor"
if linha>62 then
%>
</table>
<DIV style="page-break-after:always"></DIV>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650>
<tr>
	<td class=grupo colspan=4>RELAÇÃO DE ESTACIONAMENTO - CAMPUS <%=ucase(request.form("campus"))%></td>
	<td class=grupo colspan=2 align="right" nowrap ><%=now()%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
<tr>
	<td class=titulo>#</td>
	<td class=titulo>Código</td>
	<td class=titulo>Nome Func./Professor</td>
	<td class=titulo>Placa</td>
	<td class=titulo>Veículo</td>
	<td class=titulo>Cor</td>
</tr>
<%
linha=2
end if 'linha
%>
<tr>
	<td style="border-bottom:2 solid #000000" class=<%=classe%> rowspan=<%=rs("qt")%> align="right"><%=rs.absoluteposition%>&nbsp;</td>
	<td style="border-bottom:2 solid #000000" class=<%=classe%> rowspan=<%=rs("qt")%> ><%=rs("chapa")%></td>
	<td style="border-bottom:2 solid #000000" class=<%=classe%> rowspan=<%=rs("qt")%> ><%=rs("nome")%></td>
<%
sql2="SELECT modelo, placa, cor FROM veiculos WHERE (dttermino Is Null OR dttermino='') AND chapa='" & rs("chapa") & "' " & _
" " '& sql1a
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
do while not rs2.eof
if rs2.absoluteposition>1 then
%>
	<tr>
<%
end if
if rs2.absoluteposition=rs("qt") then estilo=" style='border-bottom:2 solid #000000'" else estilo=" "
%>
	<td class=<%=classe%> <%=estilo%> ><%=rs2("placa")%></td>
	<td class=<%=classe%> <%=estilo%> ><%=rs2("modelo")%></td>
	<td class=<%=classe%> <%=estilo%> ><%=rs2("cor")%></td>
<%
	'if rs2.recordcount>1 then response.write "&nbsp;/&nbsp;"
	linha=linha+1
rs2.movenext
%>
	</td>
	</tr>
<%
loop
end if
rs2.close
%>	
	</td>
</tr>
<%
rs.movenext
loop
rs.close
%>

<!-- Visitantes -->
<%
sql1="SELECT v.matricula, Count(placa) AS qt, nome " & _
"FROM veiculos_outros v, veiculos_outros_placas p " & _
"WHERE v.matricula=p.matricula AND v.validade='AGO/2008' AND "
if request.form("campus")="VY" then sql1=sql1 & " campus='V. Yara' "
if request.form("campus")="NS" then sql1=sql1 & " campus='Narciso' "
if request.form("campus")="JW" then sql1=sql1 & " campus='Jd.Wilson' "
sql1=sql1 & "GROUP BY v.matricula, nome " & _
" ORDER BY nome "
'response.write sql1
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<%
rs.movefirst
do while not rs.eof
'if rs("sind")="03" then classe="campoa" else classe="campot"
classe="campoa"
if linha>62 then
%>
</table>
<DIV style="page-break-after:always"></DIV>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650>
<tr>
	<td class=grupo colspan=4>RELAÇÃO DE ESTACIONAMENTO - CAMPUS <%=ucase(request.form("campus"))%></td>
	<td class=grupo colspan=2 align="right"><%=now()%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
<tr>
	<td class=titulo>#</td>
	<td class=titulo>Código</td>
	<td class=titulo>Nome Func./Professor/Visitante</td>
	<td class=titulo>Placa</td>
	<td class=titulo>Veículo</td>
	<td class=titulo>Cor</td>
</tr>
<%
linha=2
end if 'linha
%>
<tr>
	<td class=<%=classe%> rowspan=<%=rs("qt")%> align="right"><%=rs.absoluteposition%>&nbsp;</td>
	<td class=<%=classe%> rowspan=<%=rs("qt")%> ><%=rs("matricula")%></td>
	<td class=<%=classe%> rowspan=<%=rs("qt")%> ><%=rs("nome")%></td>
<%
sql2="SELECT modelo, placa, cor FROM veiculos_outros_placas WHERE matricula='" & rs("matricula") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
do while not rs2.eof
if rs2.absoluteposition>1 then
%>
	<tr>
<%
end if
%>
	<td class=<%=classe%> ><%=rs2("placa")%></td>
	<td class=<%=classe%> ><%=rs2("modelo")%></td>
	<td class=<%=classe%> ><%=rs2("cor")%></td>
<%
	'if rs2.recordcount>1 then response.write "&nbsp;/&nbsp;"
	linha=linha+1
rs2.movenext
%>
	</td>
	</tr>
<%
loop
end if
rs2.close
%>	
	
	</td>
</tr>
<%
rs.movenext
loop
end if 'rs.recordcount
rs.close
%>

<!-- alunos -->
<%
sql1="SELECT matricula, Count(placa) AS qt, nome " & _
"FROM veiculos_alunos " & _
"WHERE validade='30/06/2008' AND "
if request.form("campus")="VY" then sql1=sql1 & " campus='V. Yara' "
if request.form("campus")="NS" then sql1=sql1 & " campus='Narciso' "
if request.form("campus")="JW" then sql1=sql1 & " campus='Jd.Wilson' "
sql1=sql1 & "GROUP BY matricula, nome " & _
" ORDER BY nome "
'response.write sql
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
'if rs("sind")="03" then classe="campoa" else classe="campot"
classe="campoa"
if linha>54 then
%>
</table>
<DIV style="page-break-after:always"></DIV>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650>
<tr>
	<td class=grupo colspan=4>RELAÇÃO DE ESTACIONAMENTO - CAMPUS <%=ucase(request.form("campus"))%></td>
	<td class=grupo colspan=2 align="right"><%=now()%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
<tr>
	<td class=titulo>#</td>
	<td class=titulo>Código</td>
	<td class=titulo>Nome Func./Professor/Visitante</td>
	<td class=titulo>Placa</td>
	<td class=titulo>Veículo</td>
	<td class=titulo>Cor</td>
</tr>
<%
linha=2
end if 'linha
%>
<tr>
	<td class=<%=classe%> rowspan=<%=rs("qt")%> align="right"><%=rs.absoluteposition%>&nbsp;</td>
	<td class=<%=classe%> rowspan=<%=rs("qt")%> ><%=rs("matricula")%></td>
	<td class=<%=classe%> rowspan=<%=rs("qt")%> ><%=rs("nome")%></td>
<%
sql2="SELECT modelo, placa, cor FROM veiculos_alunos WHERE matricula='" & rs("matricula") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
do while not rs2.eof
if rs2.absoluteposition>1 then
%>
	<tr>
<%
end if
%>
	<td class=<%=classe%> ><%=rs2("placa")%></td>
	<td class=<%=classe%> ><%=rs2("modelo")%></td>
	<td class=<%=classe%> ><%=rs2("cor")%></td>
<%
	'if rs2.recordcount>1 then response.write "&nbsp;/&nbsp;"
	linha=linha+1
rs2.movenext
%>
	</td>
	</tr>
<%
loop
end if
rs2.close
%>	
	
	</td>
</tr>
<%
rs.movenext
loop
end if 'rs.recordcount>0
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