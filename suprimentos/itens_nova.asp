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
<title>Inclusão de Item de Uniforme</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
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
		if request.form("qt_novo")="" then qt_novo=0 else qt_novo=request.form("qt_novo")
		if request.form("qt_usado")="" then qt_usado=0 else qt_usado=request.form("qt_usado")
		if request.form("preco")="" then preco=0 else preco=request.form("preco")

		sql = "INSERT INTO uniforme_item (descricao, codigorm, tamanho, sequencia, qt_novo, qt_usado, preco, usuarioc, datac) "

		sql2 = " SELECT '" & request.form("descricao") & "', '" & request.form("codigorm") & "', "
		sql2=sql2 & " '" & request.form("tamanho") & "', " & request.form("sequencia") & ", "
		sql2=sql2 & " " & qt_novo & ", " & qt_usado & ", " & nraccess(preco)
		sql2=sql2 & ",'" & session("usuariomaster") & "'"
		sql2=sql2 & ",getdate()"
		sql1 = sql & sql2 & ""
		'response.write "<font size='1'>" & sql1
		if tudook=1 then conexao.Execute sql1, , adCmdText
		
		sqlr="select id_item from uniforme_item where descricao='" & request.form("descricao") & "' and tamanho='" & request.form("tamanho") & "' " & _
		"and sequencia=" & request.form("sequencia") & " and qt_novo=" & qt_novo & " and qt_usado=" & qt_usado & " and preco=" & nraccess(preco)
		rs.Open sqlr, ,adOpenStatic, adLockReadOnly
		if rs.recordcount>0 then id_item=rs("id_item")
		rs.close
		
		iNao=request.form("CNao"):'response.write "<br>" & inao
		for iLoop=0 to iNao
			id_cat=request.form("catnao" & iLoop)
			'response.write "<br>Sim " & id_cat
			if request.form("catnao" & iloop)<>"" then
				strSql="insert into uniforme_link (id_cat, id_item) values (" & id_cat & "," & id_item & ")"
				'response.write "<br>" & strsql
				conexao.execute strSql, , adCmdText
			end if
		next

	end if
else 'request.form=""
end if

if request.form="" or request.form("bt_salvar")="" then
%>
<form method="POST" action="itens_nova.asp" name="form" >
<table border="1" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo width=320 valign=top>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr><td class=grupo>Inclusão</td></tr></table>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr><td class=titulo>Cód.</td>
	<td class=titulo>Descrição</td>
</tr>
<tr><td class=titulo>0</td>
	<td class=fundo><input class=a type="text" name="descricao" size="50" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr><td class=titulo>Tamanho</td>
	<td class=titulo>Seq.</td>
	<td class=titulo>Código RM</td>
</tr>
<tr><td class=fundo><input class=a type="text" name="tamanho" size="4" value=""></td>
	<td class=fundo><input class=a type="text" name="sequencia" size="3" value=""></td>
	<td class=fundo><input class=a type="text" name="codigorm" size="15" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr><td class=titulo>Quant.Novo</td>
	<td class=titulo>Quant.Usado</td>
	<td class=titulo>Preço</td>
</tr>
<tr><td class=fundo><input class=a type="text" name="qt_novo" size="4" value=""></td>
	<td class=fundo><input class=a type="text" name="qt_usado" size="4" value=""></td>
	<td class=fundo><input class=a type="text" name="preco" size="10" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>

</td>
<td class=fundo width=180 valign=top>

<table border="0" cellpadding="3" cellspacing="0" width="180">
<tr><td class=grupo>Utilizado em:</td></tr>
<tr><td class="campor">
<%
tsim=0
tsim=tsim+1
%>
</td></tr>
<tr><td class=grupo>Não Utilizado em:</td></tr>
<tr><td class="campor">
<%
tnao=0
sqlc="select c.id_cat, c.descricao from uniforme_categoria c where id_cat<>8 order by c.descricao"
rs2.Open sqlc, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
<p style="margin-top:0;margin-bottom:0;font-size:7pt"><input type="checkbox" name="catnao<%=tnao%>" value="<%=rs2("id_cat")%>"><%=rs2("descricao")%>
<%
tnao=tnao+1:rs2.movenext:loop
rs2.close
%>
</td></tr></table>

</td></tr></table>
<input type="hidden" name="csim" value="<%=tsim-1%>">
<input type="hidden" name="cnao" value="<%=tnao-1%>">

</form>
<%
else
'rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if

'	Response.write "<p>Registro salvo.<br>"
	'response.write '<script>javascript:top.window.close();</script>
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<!--
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

set rsc=nothing
set rsd=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>