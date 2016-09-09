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
<title>Inclusão de Formação</title>
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
		sql = "INSERT INTO uprofformacao_ (" 
		sql = sql & "codprof, codinstrucao, curso, abrangencia, instituicao, localinst, anoconclusao, dataconclusao, "
		sql = sql & "comprovante, usuarioa, dataa "
		sql = sql & ") "
		sql2 = " SELECT "
		sql2=sql2 & " '" & request.form("chapa") & "', "
		sql2=sql2 & " '" & request.form("codinstrucao") & "', "
		sql2=sql2 & " '" & request.form("curso") & "', "
		sql2=sql2 & " '" & request.form("abrangencia") & "', "
		sql2=sql2 & " '" & request.form("instituicao") & "', "
		sql2=sql2 & " '" & request.form("localinst") & "', "
		sql2=sql2 & " '" & request.form("anoconclusao") & "', "
		if request.form("dataconclusao")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dataconclusao")) & "', "

		sql2=sql2 & " '" & request.form("comprovante") & "', "
		sql2=sql2 & " '" & session("usuariomaster") & "', "
		sql2=sql2 & " getdate() "
		sql1 = sql & sql2 & ""
		'response.write "<font size='2'>" & sql1
		conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
end if

if request.form="" then
%>
<form method="POST" action="formacao_nova.asp" name="planodesaude" >
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Formação Acadêmica</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=fundo><select size="1" name="chapa" class=a>
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where chapa='" & request("chapa") & "'" 
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if request("chapa")=rs2("chapa") then tempc="selected" else tempc=""
%>
          <option value="<%=rs2("chapa")%>" <%=tempc%>><%=rs2("nome")%></option>
<%
rs2.movenext:loop:rs2.close
%>
	</select></td>
</tr>
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
%>
		<option value="<%=rsc("codinstrucao")%>" <%=tempt%>><%=rsc("tipo")%></option>
<%
	rsc.movenext:loop:rsc.close
%>
	</select></td>
	<td class=fundo><input type="text" name="curso" size="25" value=""></td>
	<td class=fundo><select size="1" name="abrangencia">
		<option value="">Selecione o tipo</option>
<%
	sqla="SELECT abrangencia, descricao from uprof_abrangencia order by descricao "
	rsc.Open sqla, ,adOpenStatic, adLockReadOnly
	rsc.movefirst:do while not rsc.eof
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
	<td class=fundo><input type="text" name="instituicao" size="70" value="">  </td>
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
	<td class=fundo><input type="text" name="localinst" size="20" value=""></td>
	<td class=fundo><input type="text" name="anoconclusao" size="8" maxlength="4" value=""></td>
	<td class=fundo><input type="text" name="dataconclusao" size="8" value=""></td>
	<td class=fundo><select size="1" name="comprovante">
		<option value="N">Não</option>
		<option value="S">Sim</option>
	</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
	<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
	<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
	<input type="button" value="Fechar" class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
else
'rs.close
end if
%>
<%
if request.form("bt_salvar")<>"" then
	Response.write "<p>Registro salvo.<br>"
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
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
'conexao2.close
'set conexao2=nothing
%>
</body>
</html>