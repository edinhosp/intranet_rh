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
<title>Alteração de Lançamento em Folha</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1() { form.nome.value=form.nome.value.toUpperCase() }
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		if request.form("bolsa")="ON" then bolsa=-1 else bolsa=0
		sql="UPDATE rhconveniadosfac SET " & _
		"perlet       ='" & request.form("perlet") & "' " & _
		",nome        ='" & request.form("nome") & "' " & _
		",id_faculdade= " & request.form("faculdade") & " " & _
		",curso       ='" & request.form("curso") & "' " & _
		",periodo     ='" & request.form("periodo") & "' " & _
		",inscricao   ='" & request.form("inscricao") & "' " & _
		",class_curso = " & request.form("class_curso") & " " & _
		",class_geral = " & request.form("class_geral") & " " & _
		",pontos      = " & nraccess(request.form("pontos")) & " " & _
		",status      ='" & request.form("status") & "' " & _
		",bolsa       = " & bolsa & " " & _
		",situacao_atual='" & request.form("situacao_atual") & "' " & _
		",obs         ='" & request.form("obs") & "' " & _
		" WHERE id=" & session("id_alt_rec")
		'",usuarioa='" & session("usuariomaster") & "' "
		'",dataa   =getdate() "
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="UPDATE apont_adm set deletada=-1 WHERE id_adm=" & session("id_alt_adm")
		sql="DELETE FROM rhconveniadosfac WHERE id=" & session("id_alt_rec")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_rec=session("id_alt_rec")
		id_rec=request.form("id_rec")
	else
		id_rec=request("codigo")
	end if
	sql="select * from rhconveniadosfac where id=" & id_rec
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_rec")=rs("id")
'response.write request.form
%>
<form method="POST" action="recebidos_alteracao.asp" name="form">
<input type="hidden" name="id_rec" size="4" value="<%=rs("id")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
	<tr><td class=grupo>Alteração de Candidato <%=rs("id")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Nome do candidato</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=rs("nome")%>" name="nome" size="50" onChange="nome1()"></td>
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
if rs("id_faculdade")=rsd("id_faculdade") then tempc="selected" else tempc=""
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
if rs("periodo")=rsd("descturno") then tempc="selected" else tempc=""
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
sqla="select nome as curso, codcur from corporerm.dbo.ucursos where codcur>0 and codcur<99999 and tipocurso=2 order by nome"
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if rs("curso")=rsd("curso") then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("curso")%>" <%=tempc%>><%=rsd("curso")%></option>
<%
rsd.movenext:loop
rsd.close
%>
	</select></td>
	<td class=fundo><input type="text" value="<%=rs("perlet")%>" name="perlet" size="7"></td>
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
	<td class=fundo><input type="text" value="<%=rs("inscricao")  %>" name="inscricao" size="15"></td>
	<td class=fundo><input type="text" value="<%=rs("class_curso")%>" name="class_curso" size="5"></td>
	<td class=fundo><input type="text" value="<%=rs("class_geral")%>" name="class_geral" size="5"></td>
	<td class=fundo><input type="text" value="<%=rs("pontos")     %>" name="pontos" size="6"></td>
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
	<td class=fundo><input type="text" value="<%=rs("status")  %>" name="status" size="15"></td>
	<td class=fundo><input type="checkbox" name="bolsa" value="ON" <%if rs("bolsa")="-1" then response.write "checked"%>>&nbsp;</td>
	<td class=fundo><input type="text" value="<%=rs("situacao_atual")%>" name="situacao_atual" size="15"></td>
</tr>
</table>

<!-- status / situacao atual -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Observação</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=rs("obs")%>" name="obs" size="70"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser alterado!');</script>"
	end if
	'Response.write "<p>Registro atualizado.<br>"
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<!--
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>