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
<title>Inclusão de Uniforme - Funcionário</title>
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
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
		if request.form("bt_salvar")<>"" then
		tudook=1
		'if request.form("salvar")="1" then

if request.form("id_item")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecion o uniforme!');</script>"
end if

		sqla = "INSERT INTO uniforme_func_item (id_fcat, chapa, id_item, usuarioc, datac ) "
		sqlb = " SELECT " & request.form("id_fcat") & ""
		sqlb=sqlb & ",'" & request.form("chapa") & "' "
		sqlb=sqlb & ", " & request.form("id_item") & " "
		sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		sqlb=sqlb & ",getdate()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
		'end if
		end if 'request btsalvar
	else 'request.form=""
	end if

'if request.form("bt_salvar")="" then
if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then

if request("id_cat")<>"" then
	id_cat=request("id_cat")
elseif request.form("id_cat")<>"" then
	id_cat=request.form("id_cat") 
else
	id_cat=""
end if
if request("id_fcat")<>"" then
	id_fcat=request("id_fcat")
elseif request.form("id_fcat")<>"" then
	id_fcat=request.form("id_fcat") 
else
	id_fcat=""
end if
if request("chapa")<>"" then
	chapa=request("chapa")
elseif request.form("chapa")<>"" then
	chapa=request.form("chapa") 
else
	chapa=""
end if

%>
<form method="POST" action="func_item_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<input type="hidden" name="chapa" value="<%=chapa%>">
<input type="hidden" name="id_cat" value="<%=id_cat%>">
<input type="hidden" name="id_fcat" value="<%=id_fcat%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Inclusão de Uniforme de Funcionário</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
if request.form("id_item")="" then id_item="0" else id_item=request.form("id_item")
'if request.form("inicio")="" then inicio=formatdatetime(now,2) else inicio=request.form("inicio")
'if request.form("chapa")="" then chapa="" else chapa=request.form("chapa")
%>
<!-- uniforme -->

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Uniforme</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="id_item" onfocus="javascript:window.status='Selecione o uniforme'" >
<%
sql1="select i.id_item, i.descricao, i.tamanho from uniforme_item i, uniforme_link l " & _
"where l.id_item=i.id_item and l.id_cat=" & id_cat & " order by i.descricao, i.sequencia "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o uniforme....</option>"
rs.movefirst:do while not rs.eof
if cstr(id_item)=cstr(rs("id_item")) then temp="selected" else temp=""
%>
	<option value="<%=rs("id_item")%>" <%=temp%>><%=rs("descricao") & " - " & rs("tamanho")%></option>
<%
rs.movenext:loop:rs.close
%>
	</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
	</td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2" onfocus="javascript:window.status='Clique para desfazer e limpar a tela'"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()" onfocus="javascript:window.status='Clique aqui para fechar sem salvar'"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing
end if   'request.form=""
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<!--
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

%>
</body>
</html>