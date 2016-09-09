<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a88")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Tempo de Serviço</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql="UPDATE pfunc_compl SET "

		if request.form("data_tempo")<>"" then 
			sql=sql & "data_tempo='" & dtaccess(request.form("data_tempo")) & "', "
		else
			sql=sql & "data_tempo=null, "
		end if

		sql=sql & "tempo_trabalho= '" & request.form("tempo_trabalho")& "', "
		sql=sql & "tempo_restante= '" & request.form("tempo_restante")   & "', "

		sql=sql & "usuariot='" & session("usuariomaster") & "', "
		sql=sql & "datat   =getdate() "
		sql=sql & " WHERE chapa='" & session("alt_chapa") & "' "
		response.write sql
		conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		'sql="DELETE id_form FROM uprofformacao_ WHERE id_form=" & session("id_alt_form")
		'conexao.Execute sql, , adCmdText
	end if

else 'request.form=""

	if request("codigo")=null then
		id_form=session("alt_chapa")
	else
		id_form=request("codigo")
	end if
	sql1="select * from pfunc_compl where chapa='" & id_form & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("alt_chapa")=rs("chapa")

sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
%>
<form method="POST" action="aposentadoria_alteracao.asp" name="form">
<input type="hidden" name="id_form" size="4" value="<%=rs("chapa")%>" style="font-size: 8 pt">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Alteração de Tempo de Serviço</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=fundo><p class=realce><%=rs("chapa")%> - <%=rsnome("nome")%></p></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Dt.Informação</td>
	<td class=titulo>Tempo de Trabalho</td>
	<td class=titulo>Tempo Restante</td>
</tr>
<tr>
<%if rs("data_tempo")<>"" then data_tempo=rs("data_tempo") else data_tempo=formatdatetime(now(),2)%>
	<td class=fundo><input type="text" name="data_tempo" size="10" value="<%=data_tempo%>"></td>
	<td class=fundo><input type="text" name="tempo_trabalho" size="10" value="<%=rs("tempo_trabalho")%>"></td>
	<td class=fundo><input type="text" name="tempo_restante" size="10" value="<%=rs("tempo_restante")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
	<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
	<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
<!--	<input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td> -->
	</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if
%>
<%
if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	Response.write "Registro atualizado.<br>"
	'response.write "<a href='javascript:window.close()'>Fechar Janela</a>"
%>
 <script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
<%
end if
%>
<%
set rsc=nothing
set rsnome=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>