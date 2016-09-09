<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
	if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
	'if session("a81")="N" or session("a81")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Lista de Atividades</title>
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

if request.form("bt_salvar")<>"" then
	tudook=1
	sql="UPDATE iCRM_Fluxo SET " & _
	"anotacao = '" & request.form("anotacao") & "', " & _
	"status   = '" & request.form("status")   & "', "
	if request.form("dtvencimento")<>"" then 
		sql=sql & "dtvencimento = '" & dtaccess(request.form("dtvencimento")) & "', "
	else
		sql=sql & "dtvencimento = null, "
	end if
	sql=sql & "update_user = '" & session("usuariomaster") & "', " & _
	"update_data = '" & dtaccess(now()) & "' " & _
	"WHERE idFluxo=" & session("idFluxo")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM iCRM_Fluxo WHERE idFluxo=" & session("idFluxo")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		idFluxo=session("idFluxo")
	else
		idFluxo=request("codigo")
	end if
	sql1="select * from iCRM_Lista where idFluxo=" & idFluxo
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("idFluxo")=rs("idFluxo")

tamanho=315
%>
<form method="POST" action="lista_alteracao.asp" name="listaCRM">
<input type="hidden" name="idFluxo" size="4" value="<%=rs("idFluxo")%>">
<table border="0" cellpadding="3" cellspacing="0" width="<%=tamanho%>">
<tr><td class=grupo>Atividades</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=tamanho%>">
<tr><td class=titulo>Atividade</td>
	<td class=titulo>Tarefa</td></tr>
<tr><td class=fundo><font color=blue><%=rs("atividade")%></td>
	<td class=fundo><font color=blue><%=rs("tarefa")%></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=tamanho%>">
<tr><td class=titulo>Funcionário</td></tr>
<tr>
	<td class=fundo><%=rs("chapa")%> - <%=rs("nome")%></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=tamanho%>">
<tr><td class=titulo>Anotação</td></tr>
<tr>
	<td class=titulo><input type="text" name="anotacao" size="50" maxsize=40 value="<%=rs("anotacao")%>"> </td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=tamanho%>">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Vencimento</td>
	<td class=titulo>Status</td>
</tr>
<tr>
	<td class=titulo><%=rs("dtfluxo")%></td>
	<td class=titulo><input type="text" name="dtvencimento" value="<%=rs("dtvencimento")%>" size=10></td>
	<td class=titulo><select name="status" size="1"><option value="">Selecione...</option>
<%
sql2="select status, descricao from iCRM_Status order by ordem"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if rs2("status")=rs("status") then textosel="selected" else textosel=""
%>	
	<option value="<%=rs2("status")%>" <%=textosel%> ><%=rs2("descricao")%></option>
<%
rs2.movenext:loop
rs2.close
%>
	</select>
	</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=tamanho%>">
<tr>
	<td class=titulo align="left" >
		<input type="submit" value="Salvar" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="submit" value="Excluir" class="button" name="Bt_excluir"></td>
</tr>
</table>
<%
if rs("create_user")<>"" then txtRegC="criado por " & rs("create_user") else txtRegC=""
if rs("update_user")<>"" then txtRegU="alterado por " & rs("update_user") else txtRegU=""
if txtRegC<>"" and txtRegU<>"" then txtRegU=" e " & txtRegU
%>
<p style="margin-top:0px;margin-bottom:0px;font-size:8pt"><i>Registro <%=txtRegC%> <%=txtRegU%></i>
</form>
<%
rs.close
set rs=nothing
end if
set rs2=nothing
conexao.close
set conexao=nothing

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
%>
</body>
</html>