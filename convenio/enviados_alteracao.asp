<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a79")="N" or session("a79")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Convênio de Bolsas</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
function evento1()	{	form.codevento.value=form.evento.value;	form.submit();	}
function codigo1()	{	form.evento.value=form.codevento.value;	form.submit();	}
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
		sql="UPDATE rhconveniados SET " & _
		"id_faculdade   = " & request.form("id_faculdade") & " " & _
		",curso         ='" & request.form("curso") & "' " & _
		",periodo       ='" & request.form("periodo") & "' " & _ 
		",encaminhamento='" & request.form("encaminhamento") & "' " & _
		",obs           ='" & request.form("obs") & "' " & _
		",anoletivo     ='" & request.form("anoletivo") & "' "
		'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		'sql=sql & ",dataa   =getdate() "
		sql=sql & " WHERE id_env=" & session("id_alt_env")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM rhconveniados WHERE id_env=" & session("id_alt_env")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_env=session("id_alt_env")
		id_env=request.form("id_env")
	else
		id_env=request("codigo")
	end if
	sql="select * from rhconveniados where id_env=" & id_env
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_env")=rs("id_env")
sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
'response.write request.form
%>
<form method="POST" action="enviados_alteracao.asp" name="form">
<input type="hidden" name="id_env" size="4" value="<%=rs("id_env")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
	<tr><td class=grupo>Alteração de Convênio de Bolsa <%=rs("id_env")%></td></tr>
</table>

<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
	<td class=fundo><%=rs("id_env")%></td>
	<td class=fundo><%=rs("chapa")%></td>
	<td class=fundo><%=rsnome("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Faculdade</td>
</tr>
<tr>
<%if request.form("id_faculdade")<>"" then id_faculdade=request.form("id_faculdade") else id_faculdade=rs("id_faculdade")%>
	<td class=fundo>
		<select size="1" name="id_faculdade" onchange="javascript:submit()">
		<option value="0">Selecione uma faculdade</option>
<%
sql2="select id_faculdade, faculdade from rhconveniobe order by faculdade "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if cint(id_faculdade)=cint(rsc("id_faculdade")) then temp1="selected" else temp1=""
%>
		<option value="<%=rsc("id_faculdade")%>" <%=temp1%>><%=rsc("faculdade")%></option>
<%
rsc.movenext
loop
rsc.close
%>
	</select>
	</td>
</tr>	
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Curso</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="curso">
		<option value="0">Selecione um curso</option>
        <%
if request.form("curso")<>"" then curso=request.form("curso") else curso=rs("curso")
sql2="select cursos, id_curso from rhconveniobec where id_faculdade=" & id_faculdade & " order by cursos "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then 
rsc.movefirst:do while not rsc.eof
if curso=rsc("cursos") then temp1="selected" else temp1=""
%>
		<option value="<%=rsc("cursos")%>" <%=temp1%>><%=rsc("cursos")%></option>
<%
rsc.movenext:loop
end if 'recordcount
rsc.close
%>
	</select>	
	</td>
</tr>
</table>

<!-- Periodo/Tipo -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Período</td>
	<td class=titulo>Tipo</td>
</tr>
<tr>
<%
if request.form("periodo")<>"" then periodo=request.form("periodo") else periodo=rs("periodo")
if request.form("encaminhamento")<>"" then encaminhamento=request.form("encaminhamento") else encaminhamento=rs("encaminhamento")
%>
	<td class=fundo><select size="1" name="periodo">
		<option value="0">Selecione um período</option>
		<option value="Matutino"   <%if periodo="Matutino"   then response.write "selected"%>>Matutino</option>
		<option value="Vespertino" <%if periodo="Vespertino" then response.write "selected"%>>Vespertino</option>
		<option value="Noturno"    <%if periodo="Noturno"    then response.write "selected"%>>Noturno</option>
	</select>	
	</td>
	<td class=fundo><select size="1" name="encaminhamento">
		<option value="0">Selecione um tipo</option>
		<option value="1" <%if encaminhamento="1" then response.write "selected"%>>Inscrição no Vestibular</option>
		<option value="2" <%if encaminhamento="2" then response.write "selected"%>>Matrícula</option>
		<option value="3" <%if encaminhamento="3" then response.write "selected"%>>Renovação de Matrícula</option>
	</select>	
	</td>
</tr>
</table>

<!-- Periodo/Tipo -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Ano Letivo</td>
	<td class=titulo>Observações</td>
</tr>
<tr>
<%
if request.form("anoletivo")<>"" then anoletivo=request.form("anoletivo") else anoletivo=rs("anoletivo")
if request.form("obs")<>"" then obs=request.form("obs") else obs=rs("obs")
%>
	<td class=fundo valign="top"><input type="text" size="6" name="anoletivo" value="<%=anoletivo%>">
	</td>
	<td class=fundo><textarea name="obs" cols="30" rows="2"><%=obs%></textarea>
	</td>
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