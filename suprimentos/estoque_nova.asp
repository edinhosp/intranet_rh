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
<title>Inclusão de Movimento de Estoque</title>
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
		sqlb = " SELECT " & request.form("id_item") & ""
		sqlb=sqlb & ",'" & dtaccess(request.form("dt_movimento")) & "' "
		sqlb=sqlb & "," & request.form("id_mov") & " "
		sqlb=sqlb & "," & request.form("qt_novo") & " "
		sqlb=sqlb & "," & request.form("qt_usado") & " "
		sqlb=sqlb & ",'" & request.form("chapa") & "' "
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
%>
<form method="POST" action="estoque_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Inclusão de Movimento</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
if request("item")<>"" then
	id_item=request("item")
elseif request.form("id_item")<>"" then
	id_item=request.form("id_item") 
else
	id_item=""
end if
if request("chapa")<>"" then
	chapa=request("chapa")
elseif request.form("chapa")<>"" then
	chapa=request.form("chapa") 
else
	chapa=""
end if
if request.form("qt_novo")="" then qt_novo="0" else qt_novo=request.form("qt_novo")
if request.form("qt_usado")="" then qt_usado="0" else qt_usado=request.form("qt_usado")
if request.form("dt_movimento")="" then dt_movimento=formatdatetime(now,2) else dt_movimento=request.form("dt_movimento")
'if request.form("chapa")="" then chapa="" else chapa=request.form("chapa")
%>
<!-- uniforme -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Uniforme</td>
</tr>
<tr>
	<td class=fundo>0</td>
	<td class=fundo>
		<select size="1" name="id_item" onfocus="javascript:window.status='Selecione o tipo de uniforme'" >
<%
sql1="select id_item, descricao, tamanho from uniforme_item order by descricao, sequencia "
if id_item<>"" then sql1="SELECT i.id_item, i.descricao, i.tamanho FROM uniforme_item AS i " & _
	"WHERE i.id_item=" & id_item & " ORDER BY i.descricao, i.sequencia;"
if chapa<>"" then sql1="SELECT i.id_item, i.descricao, i.tamanho " & _
	"FROM (uniforme_item AS i INNER JOIN uniforme_link AS l ON i.id_item=l.id_item) INNER JOIN uniforme_func_cat AS fc ON l.id_cat=fc.id_cat " & _
	"WHERE fc.chapa='" & chapa & "' ORDER BY i.descricao, i.sequencia;"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o uniforme....</option>"
rs.movefirst:do while not rs.eof
if cstr(id_item)=cstr(rs("id_item")) then temp="selected" else temp=""
%>
		<option value="<%=rs("id_item")%>" <%=temp%>><%=rs("descricao") & " (" & rs("tamanho") & ")"%></option>
<%
rs.movenext:loop:rs.close
%>
		</select></td>
</tr>
</table>

<!-- data / tipo -->
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
rs.Open sql1, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o tipo....</option>"
rs.movefirst:do while not rs.eof
if cstr(request.form("id_mov"))=cstr(rs("id_mov")) then temp="selected" else temp=""
%>
		<option value="<%=rs("id_mov")%>" <%=temp%>><%=rs("descricao") & " (" & rs("tp") & ")"%></option>
<%
rs.movenext:loop:rs.close
%>
		</select></td>
	<td class=fundo><input type="text" value="<%=qt_novo%>" name="qt_novo" size="3"></td>
	<td class=fundo><input type="text" value="<%=qt_usado%>" name="qt_usado" size="3"></td>
</tr>
</table>

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
'if session("dp_chapa")<>"" then sql2=sql2 & "and chapa='" & session("dp_chapa") & "'" else sql2=sql2 & "order by nome"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o funcionário....</option>"
rs.movefirst:do while not rs.eof
if chapa=rs("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rs("chapa")%>" <%=temp%>><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
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