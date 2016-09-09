<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a5")="N" or session("a5")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Disciplinas e Professores</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
codcur=request("curso")
perlet=request("perlet")
codtur=request("turma")
curso=request("ncurso")
	'sql1="select materia from grades_materias where codmat='" & codmat & "' "
	'rs2.Open sql1, ,adOpenStatic, adLockReadOnly
	'if rs2.recordcount>0 then curso=rs2("materia") else curso=""
	'rs2.close
%>

<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr><td class=grupo colspan=8>Indice de Aprovação de Alunos - Turma <%=codtur%> - <%=curso%> - <%=perlet%></td></tr>
<tr>
	<td class=fundo align="center">Matéria</td>
	<td class=fundo align="center">Professor</td>
	<td class=fundo align="center">Aprov.</td>
	<td class=fundo align="center">Rep.Nota</td>
	<td class=fundo align="center">Rep.Freq.</td>
	<td class=fundo align="center">% Aprov</td>
</tr>
<%
sql2="select codcur, curso, perlet, codtur, g.codmat, m.materia, n_aprov, n_repnota, n_repfreq, talunos, chapa1 " & _
"FROM grades_repro g, corporerm.dbo.umaterias m " & _
"WHERE m.codmat collate database_default=g.codmat and codcur=" & codcur & " and perlet='" & perlet & "' and codtur='" & codtur & "' and talunos>0 " & _
"ORDER BY perlet, curso, m.materia "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof

sql1="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa1") & "'"
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then nome=rs2("nome") else nome=""
rs2.close

classe="campolr"
totalalunos=rs("talunos")
if cint(rs("n_aprov"))>0 then indaprov=rs("n_aprov")/totalalunos else indaprov=0
if indaprov>0.70 and indaprov<0.85 then classe="campotr"
if indaprov>0.50 and indaprov<=0.70 then classe="campoar"
if indaprov>0 and indaprov<=0.5 then classe="camporr"
%>
<tr>
	<td class=<%=classe%>><%=rs("materia")%></td>
	<td class=<%=classe%>><%=nome%></td>
	<td class=<%=classe%> align="center"><%=rs("n_aprov")%></td>
	<td class=<%=classe%> align="center"><%=rs("n_repnota")%></td>
	<td class=<%=classe%> align="center"><%=rs("n_repfreq")%></td>
	<td class=<%=classe%> align="center"><%=formatpercent(indaprov,2)%></td>
</tr>
<%
rs.movenext
loop
end if
rs.close
%>
</table>

<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>