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
<title>Inclusão de Lista de Atividades</title>
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
	if request.form("status")="" then status="A" else status=request.form("status")
	sql="insert into iCRM_Fluxo (idCRM, Chapa, DTFluxo, DtVencimento, Anotacao, Status, Create_user, Create_data) " & _
	"select '" & request.form("idCRM") & "', '" & request.form("Chapa") & "', "
	if request.form("dtfluxo")<>"" and isdate(request.form("dtFluxo")) then 
		sql=sql & "'" & dtaccess(request.form("dtFluxo")) & "', "
	else
		tudook=0
	end if
	if request.form("dtvencimento")<>"" and isdate(request.form("dtVencimento")) then 
		sql=sql & "'" & dtaccess(request.form("dtvencimento")) & "', "
	else
		tudook=0
	end if
	sql=sql & "'" & request.form("anotacao") & "', '" & status & "', " & _
	"'" & session("usuariomaster") & "',getdate() "
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

'if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
if request.form("bt_salvar")="" or (request.form("bt_salvar")<>"" and tudook=0) then
tamanho=365
%>
<form method="POST" action="lista_nova.asp" name="listaCRM">
<input type="hidden" name="idFluxo" size="4" value="<%%>">
<table border="0" cellpadding="3" cellspacing="0" width="<%=tamanho%>">
<tr><td class=grupo>Atividades</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=tamanho%>">
<tr><td class=titulo>Atividade</td>
	<td class=titulo>Tarefa</td></tr>
<tr><td class=fundo><select name="atividade" size="1" onchange="javascript:submit();"><option value="">Selecione...</option>
<%
sql2="select distinct Atividade from iCRM_Atividades order by Atividade"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if rs2("Atividade")=request.form("atividade") then textosel="selected" else textosel=""
%>	
	<option value="<%=rs2("atividade")%>" <%=textosel%> ><%=rs2("atividade")%></option>
<%
rs2.movenext:loop
rs2.close
%>
	</select>
	</td>
	<td class=fundo><select name="idCRM" size="1"><option value="">Selecione...</option>
<%
if request.form("atividade")<>"" then
sql2="select idCRM, Tarefa from iCRM_Atividades where atividade='" & request.form("atividade") & "' order by idCRM"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if rs2("idCRM")=request.form("idCRM") then textosel="selected" else textosel=""
%>	
	<option value="<%=rs2("idCRM")%>" <%=textosel%> ><%=rs2("Tarefa")%></option>
<%
rs2.movenext:loop
rs2.close
end if
%>
	</select>
	</td>
	</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=tamanho%>">
<tr><td class=titulo>Funcionário</td></tr>
<tr>
	<td class=fundo>
	<input type="text" name="chapa" size="5" value="<%=request.form("chapa")%>" onchange="javascript:submit();">
<%
if request.form("chapa")<>"" then
sql2="select chapa, nome from corporerm.dbo.Pfunc where chapa='" & request.form("chapa") & "'"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then nome=rs2("nome") else nome=""
rs2.close
end if
response.write nome
%>
	</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=tamanho%>">
<tr><td class=titulo>Anotação</td></tr>
<tr>
	<td class=titulo><input type="text" name="anotacao" size="50" maxsize=40 value="<%=anotacao%>"> </td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="<%=tamanho%>">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Vencimento</td>
	<td class=titulo>Status</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="dtfluxo" value="<%=request.form("dtfluxo")%>" size=10></td>
	<td class=titulo><input type="text" name="dtvencimento" value="<%=request.form("dtvencimento")%>" size=10></td>
	<td class=titulo><select name="status" size="1"><option value="">Selecione...</option>
<%
sql2="select status, descricao from iCRM_Status order by ordem"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if rs2("status")=request.form("status") then textosel="selected" else textosel=""
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
%>
<p style="margin-top:0px;margin-bottom:0px;font-size:8pt">
</form>
<%
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