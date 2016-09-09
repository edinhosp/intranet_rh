<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Vaga Estacionamento</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript" src="../date.js"></script>
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

if request.form("campus_trabalho")="0" or request.form("campus_estudo")="0" or request.form("periodo")="0" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe as datas do periodo aquisitivo e do gozo!');</script>"
end if
		sql="UPDATE veiculos_alunosfunc SET "
		sql=sql & " campus_trabalho='" & request.form("campus_trabalho") & "' "
		sql=sql & ",matricula      ='" & request.form("matricula") & "' "
		sql=sql & ",campus_estudo  ='" & request.form("campus_estudo") & "' "
		sql=sql & ",periodo        ='" & request.form("periodo") & "' "
		sql=sql & ",validade       ='" & request.form("validade") & "' "
		sql=sql & ",modelo         ='" & request.form("modelo") & "' "
		sql=sql & ",cor            ='" & request.form("cor") & "' "
		sql=sql & ",placa          ='" & request.form("placa") & "' "
		if request.form("dtauto")<>"" then sql=sql & ",dtauto='" & dtaccess(request.form("dtauto")) & "'"
		sql=sql & ",sequencia      ='" & request.form("sequencia") & "' "
		sql=sql & ",obs            ='" & request.form("obs") & "' "
		sql=sql & ",usuarioa       ='" & session("usuariomaster") & "' "
		sql=sql & ",dataa          = getdate() "
		sql=sql & " WHERE id_esta  =" & session("id_alt_esta")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="UPDATE apont_adm set deletada=-1 WHERE id_adm=" & session("id_alt_adm")
		sql="DELETE FROM veiculos_alunosfunc WHERE id_esta=" & session("id_alt_esta")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_esta=session("id_alt_esta")
		id_esta=request.form("id_esta")
	else
		id_esta=request("codigo")
	end if
	sql="select * from veiculos_alunosfunc where id_esta=" & id_esta
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_esta")=rs("id_esta")
sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
'response.write request.form
%>
<form method="POST" action="esta_campus_alteracao.asp" name="form">
<input type="hidden" name="id_esta" size="4" value="<%=rs("id_esta")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Alteração de Vaga no Estacionamento <%=rs("id_esta")%></td></tr>
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
	<td class=titulo><%=rs("id_esta")%></td>
	<td class=titulo><%=rs("chapa")%></td>
	<td class=titulo><%=rsnome("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Matricula</td>
	<td class=titulor>Campus<br>trabalho</td>
	<td class=titulor>Campus<br>estudo</td>
	<td class=titulor>Período<br>estudo</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="matricula" size="8" value="<%=rs("matricula")%>"></td>
<%if request.form("campus_trabalho")<>"" then tipo=request.form("campus_trabalho") else tipo=rs("campus_trabalho")%>
	<td class=fundo><select size="1" name="campus_trabalho">
		<option value="0">...</option>
		<option value="VY" <%if tipo="VY" then response.write "Selected"%>>Vila Yara</option>
		<option value="NS" <%if tipo="NS" then response.write "Selected"%>>Narciso</option>
		<option value="JW" <%if tipo="JW" then response.write "Selected"%>>Jd.Wilson</option>
		</select>
	</td>

<%if request.form("campus_estudo")<>"" then tipo=request.form("campus_estudo") else tipo=rs("campus_estudo")%>
	<td class=fundo><select size="1" name="campus_estudo">
		<option value="0">...</option>
		<option value="VY" <%if tipo="VY" then response.write "Selected"%>>Vila Yara</option>
		<option value="NS" <%if tipo="NS" then response.write "Selected"%>>Narciso</option>
		<option value="JW" <%if tipo="JW" then response.write "Selected"%>>Jd.Wilson</option>
		</select>
	</td>

<%if request.form("periodo")<>"" then tipo=request.form("periodo") else tipo=rs("periodo")%>
	<td class=fundo><select size="1" name="periodo">
		<option value="0">...</option>
		<option value="M" <%if tipo="M" then response.write "Selected"%>>Matutino</option>
		<option value="V" <%if tipo="V" then response.write "Selected"%>>Vespertino</option>
		<option value="N" <%if tipo="N" then response.write "Selected"%>>Noturno</option>
		</select>
	</td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Validade</td>
	<td class=titulo>Veículo</td>
	<td class=titulo>Cor</td>
	<td class=titulo>Placa</td>
	<td class=titulo>Dt.Autor.</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="validade" size="7" value="<%=rs("validade")%>"></td>
	<td class=fundo><input type="text" name="modelo" size="15" value="<%=rs("modelo")%>"></td>
	<td class=fundo><input type="text" name="cor" size="5" value="<%=rs("cor")%>"></td>
	<td class=fundo><input type="text" name="placa" size="7" value="<%=rs("placa")%>"></td>
	<td class=fundo><input type="text" name="dtauto" size="7" value="<%=rs("dtauto")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Sequência</td>
	<td class=titulo>Observação</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="sequencia" size="5" value="<%=rs("sequencia")%>"></td>
	<td class=fundo><input type="text" name="obs" size="45" value="<%=rs("obs")%>"></td>
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