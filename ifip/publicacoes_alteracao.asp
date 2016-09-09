<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a30")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Alteração de Publicações</title>
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

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql="UPDATE ifip_publicacoes SET "
		if request.form("dt_publicacao")=""     then
			sql=sql & "dt_publicacao=null,"
		else
			sql=sql & "dt_publicacao='" & dtaccess(request.form("dt_publicacao")) & "', "
		end if
		sql=sql & "titulo='" & request.form("titulo") & "', "
		sql=sql & "[local] ='" & request.form("local")  & "', "
		sql=sql & "numero='" & request.form("numero") & "', "
		sql=sql & "pagina='" & request.form("pagina") & "' "
		sql=sql & "WHERE id_publ=" & session("id_alt_publ")
		conexao.Execute sql ', , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		sql="DELETE FROM ifip_publicacoes WHERE id_publ=" & session("id_alt_publ")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
	if request("codigo")=null then
		id_publ=session("id_alt_publ")
	else
		id_publ=request("codigo")
	end if
	sqla="select * from ifip_publicacoes "
	sqlb="where id_publ=" & id_publ
	sql1=sqla & sqlb 
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_publ")=rs("id_publ")
%>
<form method="POST" action="publicacoes_alteracao.asp" name="form">
<input type="hidden" name="id_publ" size="4" value="<%=rs("id_publ")%>">

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Alteração de Publicação</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Data da Publicação</td>
	<td class=titulo>Título</td>
</tr>
<tr>
	<td class=fundo valign=top><input type="text" name="dt_publicacao" size="8" value="<%=rs("dt_publicacao")%>"></td>
	<td class=fundo><textarea name="titulo" cols="50" rows="3"><%=rs("titulo")%></textarea>	  </td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Local</td>
	<td class=titulo>Número</td>
	<td class=titulo>Página</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="local" size="20" value="<%=rs("local")%>"></td>
	<td class=fundo><input type="text" name="numero" size="6" value="<%=rs("numero")%>"></td>
	<td class=fundo><input type="text" name="pagina" size="6" value="<%=rs("pagina")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	Response.write "<p>Registro atualizado."
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar janela" class="button" onClick="top.window.close()">
</form>
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>