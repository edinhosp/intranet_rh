<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a56")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Candidato-Convênio IES</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {
	form.nome.value=form.nome.value.toUpperCase()
}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, ok
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1

if request.form("nome")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o nome do candidato!');</script>"
if request.form("bolsa")="ON" then bolsa=-1 else bolsa=0
		sqla = "INSERT INTO rhconveniadosfac (nome, id_faculdade, periodo, curso, perlet, inscricao, class_curso, " & _
		"class_geral, pontos, status, bolsa, situacao_atual, obs )"
		
		sqlb = " SELECT '" & request.form("nome") & "'" & _
		"," & request.form("faculdade") & "" & _
		",'" & request.form("periodo") & "'" & _
		",'" & request.form("curso") & "'" & _
		",'" & request.form("perlet") & "'" & _
		",'" & request.form("inscricao") & "'" & _
		"," & request.form("class_curso") & "" & _
		"," & request.form("class_geral") & "" & _
		"," & nraccess(request.form("pontos")) & "" & _
		",'" & request.form("status") & "'" & _
		"," & bolsa & "" & _
		",'" & request.form("situacao") & "'" & _
		",'" & request.form("obs") & "'" 
		'sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		'sqlb=sqlb & ",getdate()"
		sql = sqla & sqlb
		response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if 'request btsalvar
else 'request.form=""
end if

'if request.form("bt_salvar")="" then
%>
<form method="POST" action="recebidos_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
	<tr><td class=grupo>Inclusão de candidato</td></tr>
</table>
<!-- nome candidato -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Nome do candidato</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=request.form("nome")%>" name="nome" size="50" onChange="nome1()"></td>
</tr>
</table>

<!-- faculdade / turno -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Faculdade</td>
	<td class=titulo>Turno</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="faculdade" onfocus="javascript:window.status='Selecione o evento'">
<%
sqla="select faculdade, id_faculdade from rhconveniobe order by faculdade"
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if request.form("faculdade")=cstr(rsd("id_faculdade")) then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("id_faculdade")%>" <%=tempc%>><%=rsd("faculdade")%></option>
<%
rsd.movenext:loop
rsd.close
%>
	</select></td>
	<td class=fundo>
	<select size="1" name="periodo" onfocus="javascript:window.status='Selecione o evento'">
<%
sqla="select descturno, codturno from eturnos where codturno<>5 order by descturno"
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if request.form("periodo")=rsd("descturno") then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("descturno")%>" <%=tempc%>><%=rsd("descturno")%></option>
<%
rsd.movenext:loop
rsd.close
%>
	</select></td>
</tr>
</table>

<!-- curso -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Curso</td>
	<td class=titulo>Per.Letivo</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="curso">
<%
sqla="select nome as curso, codcur from corporerm.dbo.ucursos where codcur>0 and codcur<999 and tipocurso=2 order by nome"
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if request.form("curso")=rsd("curso") then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("curso")%>" <%=tempc%>><%=rsd("curso")%></option>
<%
rsd.movenext:loop
rsd.close
%>
	</select></td>
	<td class=fundo><input type="text" value="<%=request.form("perlet")%>" name="perlet" size="7"></td>
</tr>
</table>

<!-- matricula /classificação/pontos -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Nº Matrícula</td>
	<td class=titulo>Class.Curso</td>
	<td class=titulo>Class. na IES</td>
	<td class=titulo>Pontos</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=request.form("inscricao")  %>" name="inscricao" size="15"></td>
	<td class=fundo><input type="text" value="<%=request.form("class_curso")%>" name="class_curso" size="5"></td>
	<td class=fundo><input type="text" value="<%=request.form("class_geral")%>" name="class_geral" size="5"></td>
	<td class=fundo><input type="text" value="<%=request.form("pontos")     %>" name="pontos" size="6"></td>
</tr>
</table>

<!-- status / situacao atual -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Status</td>
	<td class=titulo>Bolsista?</td>
	<td class=titulo>Situação Atual</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=request.form("status")  %>" name="status" size="15"></td>
	<td class=fundo><input type="checkbox" name="bolsa" value="ON" <%if request.form("bolsa")="ON" then response.write "checked"%>>&nbsp;</td>
	<td class=fundo><input type="text" value="<%=request.form("situacao_atual")%>" name="situacao_atual" size="15"></td>
</tr>
</table>

<!-- status / situacao atual -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Observação</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=request.form("obs")%>" name="obs" size="70"></td>
</tr>
</table>


<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
	</td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2" onfocus="javascript:window.status='Clique para desfazer e limpar a tela'"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()" onfocus="javascript:window.status='Clique aqui para fechar sem salvar'"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing
'end if   'request.form=""
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<!--
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if
%>
</body>
</html>