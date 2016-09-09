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
<%
'exit
dim conexao,rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form("submit")="" then
%>
<form method="POST" action="carometro.asp" name="form">
Chapas: <input type="text" size=60 name="chapas" value="<%=request.form("chapas")%>" >
<br>
<input type="submit" name="submit" value="Visualizar">

</form>

<%
else 'submit

chapas=request.form("chapas")
chapas=replace(chapas,",","','")
chapas="('" & chapas & "')"


sql="select f.chapa, f.nome, f.codsecao, s.descricao, p.sexo, p.apelido, f.codsindicato " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.ppessoa p " & _
"where f.codsecao=s.codigo and f.codpessoa=p.codigo and f.chapa in " & chapas & " order by f.nome"
'response.write sql

rs.Open sql, ,adOpenStatic, adLockReadOnly
'response.write rs.recordcount
%>
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="690">
<%
coluna=1
rs.movefirst
do while not rs.eof
if coluna>5 then coluna=1
if coluna=1 then
	response.write "<tr>"
end if
%>
<tr>
	<td width=200 class="campor" valign=top>
	<IMG SRC="func_foto.asp?chapa=<%=rs("chapa")%>" width=150>
	</td>
	<td class="campop" valign=top>
	<font size=4>
	Nome: <%=rs("nome")%>
	<br>Chamado de: <%=rs("apelido")%>
	<br>Setor: <%=rs("descricao")%>
	</td>
</tr>
<%
coluna=coluna+1
if coluna>5 then
	response.write "</tr>"
end if
rs.movenext
loop
rs.close

end if

set rs=nothing
conexao.close
set conexao=nothing
%>
</table>
</BODY>
</HTML>