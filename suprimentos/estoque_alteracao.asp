<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a94")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Movimento - Estoque</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript" src="../date.js"></script>
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>
</head>
<body>
<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1

if (request.form("id_mov")="1" or request.form("id_mov")="3") and request.form("chapa")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o funcionário para esta movimentação!');</script>"
end if
if (request.form("id_mov")="") then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o tipo de movimentação!');</script>"
end if
if (request.form("id_mov")<>"1" and request.form("id_mov")<>"3") and request.form("chapa")<>"" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Este tipo de movimento não tem funcionário!');</script>"
end if

		sqla = "INSERT INTO uniforme_estoque (id_item, dt_movimento, id_mov, qt_novo, qt_usado, chapa, usuarioc, datac ) "
		sql="UPDATE uniforme_estoque SET "
		sql=sql & "id_item     =" & request.form("id_item") & " "
		sql=sql & ",dt_movimento='" & dtaccess(request.form("dt_movimento")) & "' "
		sql=sql & ",id_mov     =" & request.form("id_mov") & " "
		sql=sql & ",qt_novo    =" & request.form("qt_novo") & " "
		sql=sql & ",qt_usado   =" & request.form("qt_usado") & " "
		sql=sql & ",chapa      ='" & request.form("chapa") & "' "
		sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		sql=sql & ",dataa   =getdate() "
		sql=sql & " WHERE id_est=" & session("id_alt_estoque")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM uniforme_estoque WHERE id_est=" & session("id_alt_estoque")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_est=session("id_alt_estoque")
		id_est=request.form("id_est")
	else
		id_est=request("codigo")
	end if
	sql="select * from uniforme_estoque where id_est=" & id_est
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_estoque")=rs("id_est")
'response.write request.form

if request.form("qt_novo")="" then qt_novo=rs("qt_novo") else qt_novo=request.form("qt_novo")
if request.form("qt_usado")="" then qt_usado=rs("qt_usado") else qt_usado=request.form("qt_usado")
if request.form("dt_movimento")="" then dt_movimento=formatdatetime(rs("dt_movimento"),2) else dt_movimento=request.form("dt_movimento")
if request.form("chapa")="" then chapa=rs("chapa") else chapa=request.form("chapa")
if request.form("id_item")="" then id_item=rs("id_item") else id_item=request.form("id_item")
if request.form("id_mov")="" then id_mov=rs("id_mov") else id_mov=request.form("id_mov")

%>
<form method="POST" action="estoque_alteracao.asp" name="form">
<input type="hidden" name="id_est" size="4" value="<%=rs("id_est")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Alteração de Movimento Estoque</td></tr>
</table>

<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Uniforme</td>
</tr>
<tr>
	<td class=fundo><%=rs("id_est")%></td>
	<td class=fundo>
		<select size="1" name="id_item" onfocus="javascript:window.status='Selecione o tipo de uniforme'" >
<%
sql1="select id_item, descricao, tamanho from uniforme_item order by descricao, sequencia "
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o uniforme....</option>"
rs2.movefirst:do while not rs2.eof
if cstr(id_item)=cstr(rs2("id_item")) then temp="selected" else temp=""
%>
		<option value="<%=rs2("id_item")%>" <%=temp%>><%=rs2("descricao") & " (" & rs2("tamanho") & ")"%></option>
<%
rs2.movenext:loop:rs2.close
%>
		</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Tipo</td>
	<td class=titulo>Qt.Novo</td>
	<td class=titulo>Qt.Usado</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=dt_movimento%>" name="dt_movimento" size="9"></td>
	<td class=fundo>
		<select size="1" name="id_mov" onfocus="javascript:window.status='Selecione o tipo de movimento'" >
<%
sql1="select id_mov, descricao, tp=case when tipo=1 then 'Entrada' else 'Saida' end from uniforme_tpmov order by descricao "
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o tipo....</option>"
rs2.movefirst:do while not rs2.eof
if cstr(id_mov)=cstr(rs2("id_mov")) then temp="selected" else temp=""
%>
		<option value="<%=rs2("id_mov")%>" <%=temp%>><%=rs2("descricao") & " (" & rs2("tp") & ")"%></option>
<%
rs2.movenext:loop:rs2.close
%>
		</select></td>
	<td class=fundo><input type="text" value="<%=qt_novo%>" name="qt_novo" size="3"></td>
	<td class=fundo><input type="text" value="<%=qt_usado%>" name="qt_usado" size="3"></td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo colspan=2>Funcionário</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=chapa%>" name="chapa" size="5" onchange="chapa1()" onfocus="javascript:window.status='Informe o chapa do funcionário'"></td>
	<td class=fundo>
		<select size="1" name="nome" onchange="nome1()" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" >
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and codsindicato<>'03' order by nome "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o funcionário....</option>"
rs2.movefirst:do while not rs2.eof
if chapa=rs2("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rs2("chapa")%>" <%=temp%>><%=rs2("nome")%></option>
<%
rs2.movenext:loop
rs2.close
%>
		</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
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