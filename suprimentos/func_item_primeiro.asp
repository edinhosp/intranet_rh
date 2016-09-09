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
<title>Inclusão de Uniformes</title>
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
		'response.write request.form
		Itens=request.form("CItens")
		for iLoop=0 to Itens
			id_item=request.form("item" & iLoop)
			if request.form("item" & iloop)<>"" then
				strSql="insert into uniforme_func_item (id_fcat, chapa, id_item, usuarioc, datac ) " & _
				"values (" & request.form("id_fcat") & ",'" & request.form("chapa") & "'," & id_item & "," & _
				"'" & session("usuariomaster") & "', getdate() )"
				'response.write "<br>" & strsql
				if tudook=1 then conexao.execute strSql, , adCmdText
			end if
		next

	end if
else 'request.form=""
end if

if request.form="" or request.form("bt_salvar")="" then
id_cat=request("id_cat")
id_fcat=request("id_fcat")
chapa=request("chapa")
%>
<form method="POST" action="func_item_primeiro.asp" name="form" >
<input type="hidden" name="id_cat" value="<%=id_cat%>">
<input type="hidden" name="id_fcat" value="<%=id_fcat%>">
<input type="hidden" name="chapa" value="<%=chapa%>">
<table border="0" cellpadding="3" cellspacing="0" width="360">
<tr><td class=grupo>Inclusão</td></tr></table>

<table border="1" cellpadding="3" cellspacing="0" width="360">
<tr><td class=fundo>
<%
tnao=0
sqlc="select i.id_item, i.descricao, i.tamanho from uniforme_item i, uniforme_link l " & _
"where l.id_item=i.id_item and l.id_cat=" & id_cat & " order by i.descricao, i.sequencia "
rs2.Open sqlc, ,adOpenStatic, adLockReadOnly
total_itens=rs2.recordcount
col1=int(total_itens/2)
col2=total_itens-col1
rs2.movefirst:do while not rs2.eof

if rs2.absoluteposition=col2+1 then
	response.write "</td><td class=fundo>"
end if
%>
<p style="margin-top:0;margin-bottom:0;font-size:7pt">
<input type="checkbox" name="item<%=tnao%>" value="<%=rs2("id_item")%>"><%=rs2("descricao")%> (<%=rs2("tamanho")%>)
<%
tnao=tnao+1:rs2.movenext:loop
rs2.close
%>
</td></tr></table>

<table border="0" cellpadding="3" cellspacing="0" width="360">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
<input type="hidden" name="citens" value="<%=tnao-1%>">

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