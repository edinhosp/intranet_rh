<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a36")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Controle de Férias</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<!--<script language="JavaScript" type="text/javascript" src="../date.js"></script> -->
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
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

if request.form("dtinigozo")="" or request.form("dtfimgozo")="" or request.form("dtfimper")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe as datas do periodo aquisitivo e do gozo!');</script>"
end if
		dias=request.form("dias")
		if request.form("tipo")="P" then dias=0
		if dias<0 then dias=cdate(request.form("dtfimgozo"))-cdate(request.form("dtinigozo"))+1
		sql="UPDATE ferias SET "
		if request.form("dtfimper")<>"" then sql=sql & "dtfimper='" & dtaccess(request.form("dtfimper")) & "', "
		if request.form("dtinigozo")<>"" then sql=sql & "dtinigozo ='" & dtaccess(request.form("dtinigozo")) & "', "
		if request.form("dtfimgozo")<>"" then sql=sql & "dtfimgozo ='" & dtaccess(request.form("dtfimgozo")) & "', "
		sql=sql & "dias      = " & dias & " "
		sql=sql & ",tipo      ='" & request.form("tipo") & "' "
		sql=sql & ",chapa     ='" & request.form("chapa") & "' "
		sql=sql & ",obs       ='" & request.form("obs") & "' "
		'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		'sql=sql & ",dataa   =getdate() "
		sql=sql & " WHERE id_fer=" & session("id_alt_fer")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="UPDATE apont_adm set deletada=-1 WHERE id_adm=" & session("id_alt_adm")
		sql="DELETE FROM ferias WHERE id_fer=" & session("id_alt_fer")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_fer=session("id_alt_fer")
		id_fer=request.form("id_fer")
	else
		id_fer=request("codigo")
	end if
	sql="select * from ferias where id_fer=" & id_fer
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_fer")=rs("id_fer")
sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
'response.write request.form
%>
<form method="POST" action="fer_alteracao.asp" name="form">
<input type="hidden" name="id_fer" size="4" value="<%=rs("id_fer")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Alteração de Controle de Férias <%=rs("id_fer")%></td></tr>
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
	<td class=titulo><%=rs("id_fer")%></td>
	<td class=titulo><%=rs("chapa")%></td>
	<td class=titulo><%=rsnome("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Tipo</td>
	<td class=titulo>Venc.PAq</td>
	<td class=titulo>Obs.</td>
</tr>
<tr>
<%if request.form("tipo")<>"" then tipo=request.form("tipo") else tipo=rs("tipo")%>
	<td class=fundo><select size="1" name="tipo">
		<option value="C" <%if tipo="C" then response.write "Selected"%>>Crédito</option>
		<option value="D" <%if tipo="D" then response.write "Selected"%>>Débito</option>
		<option value="P" <%if tipo="P" then response.write "Selected"%>>Prev.Gozo</option>
		</select>
	</td>
	<td class=fundo><input type="text" name="dtfimper" size="9" value="<%=rs("dtfimper")%>" >
	</td>
	<td class=fundo><input type="text" name="obs" size="20" value="<%=rs("obs")%>"></td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<!-- onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')" -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Início Gozo</td>
	<td class=titulo>Término Gozo</td>
	<td class=titulo>Dias</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="dtinigozo" onchange="mand_ini1(1)" size="9" value="<%=rs("dtinigozo")%>" ></td>
	<td class=fundo><input type="text" name="dtfimgozo" onchange="mand_ini1(1)" size="9" value="<%=rs("dtfimgozo")%>" ></td>
	<td class=fundo><input type="text" name="dias" size="8" value="<%=rs("dias")%>" onfocus="this.blur()" ></td>
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