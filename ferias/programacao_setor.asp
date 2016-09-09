<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a42")="N" or session("a42")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Alterar Seção para Programação de férias</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql="UPDATE feriasprog SET relsecao='" & request.form("codsecao") & "' where chapa='" & session("alt_chapa") & "' " & _
		"and sessao='" & session("usuariomaster") & "' "
		response.write "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>" & sql
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
	sql1="select top 1 * from feriasprog where chapa='" & id_form & "' "
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
<form method="POST" action="programacao_setor.asp" name="form">
<input type="hidden" name="id_form" size="4" value="<%=rs("chapa")%>" style="font-size: 8 pt">
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=fundo nowrap><p class=realce><%=rs("chapa")%> - <%=rsnome("nome")%></p></td></tr>
<tr><td class=titulo>Seção para programação</td></tr>
<tr><td class=fundo nowrap>
<select name="codsecao" class=a>
<%
sqla="select codigo codsecao, descricao nome from corporerm.dbo.psecao s order by s.descricao "
rs2.Open sqla, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if rs2("codsecao")=rs("codsecao") then temps="selected" else temps=""
%>
	<option value="<%=rs2("codsecao")%>" <%=temps%> ><%=rs2("nome")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
</td></tr>

</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
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
	response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location.reload();self.close();</script>"
%>
<!-- <script language="Javascript">javascript:window.opener.location.submit</script> -->
<!-- <script language="Javascript">javascript:window.opener.location=window.opener.location</script> -->
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