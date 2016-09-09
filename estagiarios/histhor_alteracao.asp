<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Histórico de Horário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript" src="../date.js"></script>
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1()		{form.chapa.value=form.nome.value;}
function chapa1()		{form.nome.value=form.chapa.value;}
function descricao1()	{form.codigo.value=form.descricao.value;}
function codigo1()		{form.descricao.value=form.codigo.value;}

function mand_ini1(muda) {
	temp=form.dtinigozo.value;
	inicio=new Date(temp.substr(6),temp.substr(3,2)-1,temp.substr(0,2));
	temp2=form.dtfimgozo.value;
	termino=new Date(temp2.substr(6),temp2.substr(3,2)-1,temp2.substr(0,2));
	dinicio=montharray[inicio.getMonth()]+" "+inicio.getDate()+", "+inicio.getFullYear()
	dfinal=montharray[termino.getMonth()]+" "+termino.getDate()+", "+termino.getFullYear()
	dias=(Math.round((Date.parse(dfinal)-Date.parse(dinicio))/(24*60*60*1000))*1)+1
	document.form.dias.value=dias
}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1

if request.form("dtmudanca")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a data da mudança!');</script>"
if request.form("codigo")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o código do horário!');</script>"

		sql="UPDATE est_histhor SET "
		sql=sql & "dtmudanca='" & dtaccess(request.form("dtmudanca")) & "', "
		sql=sql & "codigo='" & request.form("codigo") & "', "
		sql=sql & "dia=" & request.form("dia") & " "
		'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		'sql=sql & ",dataa   =getdate() "
		sql=sql & " WHERE id_hist=" & session("id_alt_hist")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM est_histhor WHERE id_hist=" & session("id_alt_hist")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigoh")="" then
		id_hist=session("id_alt_hist")
		id_hist=request.form("id_hist")
	else
		id_hist=request("codigoh")
	end if
	sql="select * from est_histhor where id_hist=" & id_hist
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_hist")=rs("id_hist")
sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
'response.write request.form
if request.form("codigo")="" then codigo=rs("codigo") else codigo=request.form("codigo")
if request.form("dtmudanca")="" then dtmudanca=rs("dtmudanca") else dtmudanca=request.form("dtmudanca")
if request.form("dia")="" then dia=rs("dia") else dia=request.form("dia")
%>
<form method="POST" action="histhor_alteracao.asp" name="form">
<input type="hidden" name="id_hist" size="4" value="<%=rs("id_hist")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Alteração de Histórico de Horário <%=rs("id_hist")%></td></tr>
</table>

<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
	<td class=titulo><%=rs("id_hist")%></td>
	<td class=titulo><%=rs("chapa")%></td>
	<td class=titulo><%=rsnome("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Horário</td>
</tr>
<tr>
	<td class=fundo>
		<input type="text" value="<%=codigo%>" name="codigo" size="5" onchange="codigo1();form.submit();" onfocus="javascript:window.status='Informe o codigo do horário';">
		<select size="1" name="descricao" onchange="descricao1();form.submit();" onfocus="javascript:window.status='Selecione o horário';" >
<%
sql2="select codigo, descricao from est_cadhorario order by descricao "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o horário....</option>"
rsc.movefirst:do while not rsc.eof
if codigo=rsc("codigo") then temp="selected" else temp=""
%>
		<option value="<%=rsc("codigo")%>" <%=temp%>><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
rsc.close
%>
		</select></td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Data Mudança</td>
	<td class=titulo>Indice</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="dtmudanca" size="9" value="<%=dtmudanca%>"></td>
	<td class=fundo>
		<select size="1" name="dia" onchange="" onfocus="javascript:window.status='Selecione o dia de indice'" >
<%
sql2="select dia, comp, [desc] from est_cadhorario_marcacoes where codigo='" & codigo & "' order by dia "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if dia=rsc("dia") then temp="selected" else temp=""
%>
		<option value="<%=rsc("dia")%>" <%=temp%>><%=rsc("dia")%></option>
<%
rsc.movenext:loop
rsc.close
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