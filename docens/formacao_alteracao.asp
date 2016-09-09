<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a7")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Formação</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
'set conexao2=server.createobject ("ADODB.Connection")
'conexao2.Open application("consql")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		sql="UPDATE uprofformacao_ SET "
		sql=sql & "codinstrucao= '" & request.form("codinstrucao") & "', "
		sql=sql & "curso       = '" & request.form("curso")& "', "
		sql=sql & "abrangencia = '" & request.form("abrangencia")   & "', "
		sql=sql & "instituicao = '" & request.form("instituicao")   & "', "
		sql=sql & "localinst   = '" & request.form("localinst") & "', "
		sql=sql & "anoconclusao= '" & request.form("anoconclusao") & "', "
		if request.form("dataconclusao")<>"" then 
			sql=sql & "dataconclusao='" & dtaccess(request.form("dataconclusao")) & "', "
		else
			sql=sql & "dataconclusao=null, "
		end if
		sql=sql & "comprovante ='" & request.form("comprovante") & "', "
		sql=sql & "usuarioa='" & session("usuariomaster") & "', "
		sql=sql & "dataa   =getdate() "
		sql=sql & " WHERE id_form=" & session("id_alt_form")
		conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		sql="DELETE FROM uprofformacao_ WHERE id_form=" & session("id_alt_form")
		conexao.Execute sql, , adCmdText
	end if

else 'request.form=""

	if request("codigo")=null then
		id_form=session("id_alt_form")
	else
		id_form=request("codigo")
	end if
	sqla="select * from uprofformacao_ "
	sqlb="where id_form=" & id_form
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_form")=rs("id_form")

sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("codprof") & "'"
rs2.Open sqlz, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then nome=rs2("nome") else nome=""
rs2.close
%>
<form method="POST" action="formacao_alteracao.asp" name="form">
<input type="hidden" name="id_form" size="4" value="<%=rs("id_form")%>" style="font-size: 8 pt">
<input type="hidden" name="chapa" size="5" value="<%=rs("codprof")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Formação Acadêmica</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=fundo><p class=realce><%=rs("codprof")%> - <%=nome%></p></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Tipo</td>
	<td class=titulo>Curso</td>
	<td class=titulo>Abrangência</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="codinstrucao">
		<option value="">Selecione o tipo</option>
<%
	sqla="SELECT codinstrucao, tipo from uprof_tipo order by codinstrucao"
	rsc.Open sqla, ,adOpenStatic, adLockReadOnly
	rsc.movefirst:do while not rsc.eof
	if rsc("codinstrucao")=rs("codinstrucao") then tempt="selected" else tempt=""
%>
		<option value="<%=rsc("codinstrucao")%>" <%=tempt%>><%=rsc("tipo")%></option>
<%
	rsc.movenext:loop:rsc.close
%>
	</select></td>
	<td class=fundo><input type="text" name="curso" size="25" value="<%=rs("curso")%>"></td>
	<td class=fundo><select size="1" name="abrangencia">
		<option value="">Selecione o tipo</option>
<%
	sqla="SELECT abrangencia, descricao from uprof_abrangencia order by descricao "
	rsc.Open sqla, ,adOpenStatic, adLockReadOnly
	rsc.movefirst:do while not rsc.eof
	if rsc("abrangencia")=rs("abrangencia") then tempt="selected" else tempt=""
%>
		<option value="<%=rsc("abrangencia")%>" <%=tempt%>><%=rsc("descricao")%></option>
<%
	rsc.movenext:loop:rsc.close
%>
	</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Instituição</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="instituicao" size="70" value="<%=rs("instituicao")%>">  </td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Local Inst.</td>
	<td class=titulo>ANO conclusão</td>
	<td class=titulo>Data conclusão</td>
	<td class=titulo>Comprovante</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="localinst" size="20" value="<%=rs("localinst")%>"></td>
	<td class=fundo><input type="text" name="anoconclusao" size="8" value="<%=rs("anoconclusao")%>"></td>
	<td class=fundo><input type="text" name="dataconclusao" size="8" value="<%=rs("dataconclusao")%>"></td>
	<td class=fundo><select size="1" name="comprovante">
		<option value="N" <%if rs("comprovante")="N" then response.write "selected"%> >Não</option>
		<option value="S" <%if rs("comprovante")="S" then response.write "selected"%> >Sim</option>
	</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
	<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
	<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
	<input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
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
conexao.close
set conexao=nothing
'conexao2.close
'set conexao2=nothing
%></body>
</html>