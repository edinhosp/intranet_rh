<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.redirect "intranet.asp"
if session("a1")="N" or session("a1")="" then response.redirect "intranet.asp"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pesquisa Curriculos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->
<form method="POST" action="curriculo.asp" name="form">
Disciplina <input type="text" name="disciplina" size="20" class="form_box" value="<%=request.form("disciplina")%>">

<%
if request.form("disciplina")<>"" then
disciplina=request.form("disciplina")

dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("mysqlfieo")
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=localhost; Port=3306; Option=0; Socket=; Stmt=; Database=rhonline2; Uid=root; Pwd="
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=colossus2.fieo.br; Port=3306; Option=0; Socket=; Stmt=; Database=website; Uid=rh; Pwd=!@#qaz"

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

sqla="SELECT * from tb_rh_candidato limit 10"
if disciplina="todos" then
	sqla="select distinct cpf, c.nome, tipo_candidato, email, lattes, observacoes, null as cod, null as disciplina, cidade, tel_celular " & _
	"from tb_rh_candidato c where tipo_candidato=1 order by nome "
else
	sqla="select r.cpf, c.nome, tipo_candidato, email, lattes, observacoes, r.disciplina as cod, d.disciplina, d.area, cidade, tel_celular " & _
	"from tb_rh_rel_candidato_disciplina r " & _
	"inner join tb_rh_disciplina d on d.id_disciplina=r.disciplina " & _
	"inner join tb_rh_candidato c on c.cpf=r.cpf " & _
	"where d.disciplina like '%" & disciplina & "%' and tipo_candidato=1 order by nome"
end if

rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" width="800" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Disciplina</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Email</td>
	<td class=titulo>Celular</td>
	<td class=titulo>Cidade</td>
</tr>
<%
do while not rs.eof
if disciplina="todos" then campo1=rs("lattes") else campo1=rs("disciplina")
%>
<tr>
	<td class=campo><%=campo1%></td>
	<td class=campo nowrap><b><%=rs("nome")%></td>
	<td class=campo><%=rs("email")%></td>
	<td class=campo><%=rs("tel_Celular")%></td>
	<td class=campo><%=rs("cidade")%></td>
</tr>
<%if disciplina<>"todos" then%>
<tr>
	<td class=campo>&nbsp;</td>
	<td class=campo colspan=4><a href='<%=rs("lattes")%>' target=_blank><%=rs("lattes")%></a></td>
</tr>
<%end if%>
<%if rs("observacoes")<>"" then %>
<tr>
	<td class=campo style="border-bottom:2px solid #000000">&nbsp;</td>
	<td class=campo style="border-bottom:2px solid #000000" colspan=4><%=rs("observacoes")%></td>
</tr>
<%end if%>

<%
rs.movenext
loop
%>
</table>
<%
%>

<%
''*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'rs.movefirst
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
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
'*************** fim teste **********************%>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing

end if 'request.form("disciplina")<>""
%>


</form>
<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>