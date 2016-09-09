<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Display Image</TITLE>
<link rel="stylesheet" type="text/css" href="diversos.css">
</HEAD>
<BODY>
<form name="form" method="POST" action="ali.asp" >
Turma: <input type="text" name="turma" size="8" value="<%=request.form("turma")%>">
Periodo Letivo: <input type="text" name="perlet" size="6" value="<%=request.form("perlet")%>">
Sexo: <input type="text" name="sexo" size="1" value="<%=request.form("sexo")%>">
<input type="submit" name="ok" value="OK">
</form>


<%
'exit
dim conexao,rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form("turma")="" then turma="0" else turma=request.form("turma")
if request.form("perlet")="" then perlet="0" else perlet=request.form("perlet")
if request.form("sexo")="" then sexo="" else sexo=request.form("sexo")
if sexo="" then ssexo=" and e.sexo in ('F','M') " else ssexo=" and e.sexo in ('" & sexo & "') "

sql="SELECT um.MATALUNO, um.CODTUR, um.PERLETIVO, e.NOME, e.SEXO, e.DTNASC, Count(um.CODMAT) AS disc, cast(cast(getdate()-e.dtnasc as int)/365.25 as int) as idade " & _
", min(e.idimagem) as idimagem, min(um.mataluno) as matricula, e.nome as aluno, e.email " & _
"FROM corporerm.dbo.EALUNOS e INNER JOIN corporerm.dbo.UMATALUN um ON e.MATRICULA=um.MATALUNO " & _
"/* where um.codmat='g0186' */ " & _
"GROUP BY um.MATALUNO, um.CODTUR, um.PERLETIVO, e.NOME, e.SEXO, e.DTNASC, e.email " & _
"HAVING um.CODTUR LIKE '" & turma & "%' AND um.PERLETIVO like '" & perlet & "' " & ssexo & " " & _
"ORDER BY e.nome"
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then '--------------------
response.write rs.recordcount
%>
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
<%
coluna=1
rs.movefirst
do while not rs.eof
if coluna>5 then coluna=1
if coluna=1 then
	response.write "<tr>"
end if
%>
	<td width=120 class="campor" valign=top>
	<IMG SRC="alimg.asp?id=<%=rs("idimagem")%>" width="120">
	<br><%=rs("aluno")%> - <%=rs("matricula")%>
	<br><%=rs("dtnasc")%> - <%=rs("codtur")%>
	<br><%=rs("email")%>
	</td>
<%
coluna=coluna+1
if coluna>5 then
	response.write "</tr>"
end if
rs.movenext
loop
end if '------------------------------

rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</table>

</BODY>
</HTML>