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
<title>Alteração de Uniforme - Funcionário</title>
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

if request.form("id_item")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a categoria!');</script>"
end if

		sql="UPDATE uniforme_func_item SET "
		sql=sql & "id_item     =" & request.form("id_item") & " "
		sql=sql & ",chapa      ='" & request.form("chapa") & "' "
		sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		sql=sql & ",dataa   =getdate() "
		sql=sql & " WHERE id_fitem=" & session("id_alt_fitem")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM uniforme_func_item WHERE id_fitem=" & session("id_alt_fitem")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" an<u></u>d request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_fitem=session("id_alt_fitem")
		id_fitem=request.form("id_fitem")
	else
		id_fitem=request("codigo")
	end if
	sql="select * from uniforme_func_item where id_fitem=" & id_fitem
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
session("id_alt_fitem")=rs("id_fitem")

'if request.form("inicio")="" then inicio=formatdatetime(rs("inicio"),2) else inicio=request.form("inicio")
if request.form("id_item")="" then id_item=rs("id_item") else id_item=request.form("id_item")
if request.form("chapa")="" then chapa=rs("chapa") else chapa=request.form("chapa")

%>
<form method="POST" action="func_item_alteracao.asp" name="form">
<input type="hidden" name="id_fitem" size="4" value="<%=rs("id_fitem")%>">
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Alteração de Uniforme de Funcionário</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Uniforme</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="id_item" onfocus="javascript:window.status='Selecione o uniforme'" >
<%
sql1="select id_item, descricao, tamanho from uniforme_item order by descricao, sequencia "
sql1="SELECT i.id_item, i.descricao, i.tamanho " & _
"FROM (uniforme_link l INNER JOIN uniforme_func_cat fc ON l.id_cat=fc.id_cat) INNER JOIN uniforme_item i ON l.id_item=i.id_item " & _
"WHERE fc.id_fcat=" & rs("id_fcat") & " ORDER BY i.descricao, i.sequencia;"
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione a categoria....</option>"
rs2.movefirst:do while not rs2.eof
if cstr(id_item)=cstr(rs2("id_item")) then temp="selected" else temp=""
%>
	<option value="<%=rs2("id_item")%>" <%=temp%>><%=rs2("descricao") & " - " & rs2("tamanho")%></option>
<%
rs2.movenext:loop:rs2.close
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