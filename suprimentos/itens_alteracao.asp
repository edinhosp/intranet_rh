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
<title>Alteração de Item de Uniforme</title>
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
		if request.form("qt_novo")="" then qt_novo=0 else qt_novo=request.form("qt_novo")
		if request.form("qt_usado")="" then qt_usado=0 else qt_usado=request.form("qt_usado")
		if request.form("preco")="" then preco=0 else preco=request.form("preco")
		tudook=1
		sql="UPDATE uniforme_item SET "
		sql=sql & "descricao='" & request.form("descricao") & "', "
		sql=sql & "codigorm ='" & request.form("codigorm") & "', "
		sql=sql & "tamanho  ='" & request.form("tamanho") & "', "
		sql=sql & "sequencia=" & request.form("sequencia") & ", "
		sql=sql & "qt_novo= " & qt_novo  & ", "
		sql=sql & "qt_usado=" & qt_usado & ", "
		sql=sql & "preco=   " & nraccess(preco) & " "
		sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		sql=sql & ",dataa   =getdate() "
		sql=sql & "WHERE id_item=" & session("id_alt_item")
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText

		iSim=request.form("CSim"):'response.write "<br>" & isim
		iNao=request.form("CNao"):'response.write "<br>" & inao
		for iLoop=0 to iSim
			id_cat=request.form("catsim" & iLoop)
			'response.write "<br>Nao " & id_cat
			if request.form("catsim" & iloop)<>"" then
				strSql="delete from uniforme_link where id_cat=" & id_cat & " and id_item=" & session("id_alt_item")
				'response.write "<br>" & strsql
				conexao.execute strSql, , adCmdText
			end if
		next
		for iLoop=0 to iNao
			id_cat=request.form("catnao" & iLoop)
			'response.write "<br>Sim " & id_cat
			if request.form("catnao" & iloop)<>"" then
				strSql="insert into uniforme_link (id_cat, id_item) values (" & id_cat & "," & session("id_alt_item") & ")"
				'response.write "<br>" & strsql
				conexao.execute strSql, , adCmdText
			end if
		next

	end if 'bt_salvar
	
	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM uniforme_item WHERE id_item=" & session("id_alt_item")
		conexao.Execute sql, , adCmdText
	end if

else 'request.form=""

	if request("codigo")=null then
		id_item=session("id_alt_item")
	else
		id_item=request("codigo")
	end if
	sqla="select * from uniforme_item where id_item=" & id_item	
	rs.Open sqla, ,adOpenStatic, adLockReadOnly

end if
%>

<%
if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_item")=rs("id_item")
%>
<form method="POST" action="itens_alteracao.asp" name="form">
<input type="hidden" name="id_item" size="4" value="<%=rs("id_item")%>">  

<table border="1" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo width=320 valign=top>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr><td class=grupo>Alteração</td></tr></table>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr><td class=titulo>Cód.</td>
	<td class=titulo>Descrição</td>
</tr>
<tr><td class=titulo><%=rs("id_item")%></td>
	<td class=fundo><input class=a type="text" name="descricao" size="50" value="<%=rs("descricao")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr><td class=titulo>Tamanho</td>
	<td class=titulo>Seq.</td>
	<td class=titulo>Código RM</td>
</tr>
<tr><td class=fundo><input class=a type="text" name="tamanho" size="4" value="<%=rs("tamanho")%>"></td>
	<td class=fundo><input class=a type="text" name="sequencia" size="3" value="<%=rs("sequencia")%>"></td>
	<td class=fundo><input class=a type="text" name="codigorm" size="15" value="<%=rs("codigorm")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr><td class=titulo>Quant.Novo</td>
	<td class=titulo>Quant.Usado</td>
	<td class=titulo>Preço</td>
</tr>
<tr><td class=fundo><input class=a type="text" name="qt_novo" size="4" value="<%=rs("qt_novo")%>"></td>
	<td class=fundo><input class=a type="text" name="qt_usado" size="4" value="<%=rs("qt_usado")%>"></td>
	<td class=fundo><input class=a type="text" name="preco" size="10" value="<%=rs("preco")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="320">
<tr>
	<td class=titulo align="center" rowspan=2><input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
</tr>
<tr>
	<td class=titulo align="center"><input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
</tr>
</table>

</td>
<td class=fundo width=180 valign=top>

<table border="0" cellpadding="3" cellspacing="0" width="180">
<tr><td class=grupo>Utilizado em:</td></tr>
<tr><td class="campor">
<%
tsim=0
sqlc="select c.id_cat, c.descricao from uniforme_categoria c, uniforme_link l where l.id_cat=c.id_cat and l.id_item=" & rs("id_item") & " order by c.descricao"
rs2.Open sqlc, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
<p style="margin-top:0;margin-bottom:0;font-size:7pt"><input type="checkbox" name="catsim<%=tsim%>" value="<%=rs2("id_cat")%>"><%=rs2("descricao")%>
<%
tsim=tsim+1:rs2.movenext:loop
end if
rs2.close
%>
</td></tr>
<tr><td class=grupo>Não Utilizado em:</td></tr>
<tr><td class="campor">
<%
tnao=0
sqlc="select c.id_cat, c.descricao from uniforme_categoria c where c.id_cat not in (select id_cat from uniforme_link where id_item=" & rs("id_item") & ") and id_cat<>8 order by c.descricao"
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

set rsc=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>