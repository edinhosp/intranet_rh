<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a83")="N" or session("a83")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Lançamento VT</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>
<script language="VBScript">
	Sub tarifa_onChange
		ok=false:dim frm:set frm=document.form
		temp=cdbl(document.form.tarifa.value)
		temp2=cint(document.form.quantidade.value)
		calculo=temp*temp2
		document.form.total.value=calculo
	End Sub
	Sub quantidade_onChange
		ok=false:dim frm:set frm=document.form
		temp=cdbl(document.form.tarifa.value)
		temp2=cint(document.form.quantidade.value)
		calculo=temp*temp2
		document.form.total.value=calculo
	End Sub
</script>
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

if request.form("bt_salvar")<>"" then
	tudook=1
	total=cdbl(request.form("tarifa"))*cint(request.form("quantidade"))
	
	sql="UPDATE vt_saldo SET "
	if request.form("data")="" then sql=sql & "data=null" else sql=sql & "data='" & dtaccess(request.form("data")) & "' "
	sql=sql & ",id_tipo   =" & request.form("id_tipo") & " "
	sql=sql & ",codigo    ='" & request.form("codigop") & "' "
	sql=sql & ",tarifa    =" & nraccess(request.form("tarifa")) & " "
	sql=sql & ",quantidade=" & nraccess(request.form("quantidade")) & " "
	sql=sql & ",total     =" & nraccess(total) & " "
	sql=sql & ",chapa     ='" & request.form("chapa") & "' "
	'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
	'sql=sql & ",dataa   =now() "
	sql=sql & "WHERE id_saldo=" & session("id_alt_saldo")
	'response.write sql
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	'sql="DELETE id_grade FROM grades WHERE id_grade=" & session("id_alt_grade")
	sql="UPDATE vt_saldo set deletada=1 WHERE id_saldo=" & session("id_alt_saldo")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_saldo=session("id_alt_saldo")
		id_saldo=request.form("id_saldo")
	else
		id_saldo=request("codigo")
	end if
	sql="select * from vt_saldo where id_saldo=" & id_saldo
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_saldo")=rs("id_saldo")
'response.write request.form
%>
<form method="POST" action="mov_alteracao.asp" name="form">
<input type="hidden" name="id_saldo" size="4" value="<%=rs("id_saldo")%>">  
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Alteração de Lançamento VT <%=rs("id_saldo")%></td></tr>
</table>

<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Tipo Movimento</td>
	<td class=titulo>Vale Transporte</td>
</tr>
<tr>
<%if request.form("data")<>"" then dataf=request.form("data") else dataf=rs("data")%>
	<td class=titulo><input type="text" name="data" size="12" value="<%=formatdatetime(dataf,2)%>"></td>
	<td class=titulo><select size="1" name="id_tipo" onfocus="javascript:window.status='Selecione o tipo de movimento'">
<%
if request.form("id_tipo")<>"" then id_tipo=request.form("id_tipo") else id_tipo=rs("id_tipo")
sqla="SELECT * from vt_saldo_tipo where id_tipo<>0 "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>....</option>"
rsd.movefirst:do while not rsd.eof
if cint(id_tipo)=cint(rsd("id_tipo")) then temppl="selected" else temppl=""
%>
		<option value="<%=rsd("id_tipo")%>" <%=temppl%>><%=rsd("tipo")%></option>
<%
rsd.movenext:loop
rsd.close
%>
	</select></td>
	
	<td class=fundo><select size="1" name="codigop" onChange="javascript:submit()" onfocus="javascript:window.status='Selecione o Vale Transporte'">
<%
if request.form("codigop")<>"" then codigop=request.form("codigop") else codigop=rs("codigo")
sqla="select codigo, descricao from corporerm.dbo.ptarifa where getdate() between iniciovigencia and finalvigencia and codigo in ('01','02','03','04','06','23','DRE','S01','S02','S03','S04','S05','S06','SPC','SPI','SPM') order by descricao "

rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if codigop=rsd("codigo") then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("codigo")%>" <%=tempc%>><%=rsd("descricao")%></option>
<%
rsd.movenext:loop
rsd.close
%>
	</select></td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Tarifa</td>
	<td class=titulo>Quantidade</td>
	<td class=titulo>Total</td>
</tr>
<tr>
<%
sql="select valor from corporerm.dbo.ptarifa where codigo='" & codigop & "' and '" & dtaccess(request.form("data")) & "' between iniciovigencia and finalvigencia "
rsd.Open sql, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then tarifa=rsd("valor") else tarifa=0
rsd.close
if request.form("quantidade")<>"" then quantidade=request.form("quantidade") else quantidade=rs("quantidade")
if request.form("total")<>"" then total=request.form("total") else total=rs("total")
total=cdbl(tarifa)*cint(quantidade)
%>
<input type="hidden" name="totalt" value="<%=total%>">
	<td class=fundo><input type="text" name="tarifa"     class=vr size="8" value="<%=formatnumber(tarifa,2)%>"></td>
	<td class=fundo><input type="text" name="quantidade" class=vr size="8" value="<%=formatnumber(quantidade,0)%>"></td>
	<td class=fundo><input type="text" name="total"      class=vr size="8" value="<%=formatnumber(total,2)%>" disabled></td>
</tr>
</table>

<!-- Chapa / Nome -->
<%
if request.form("chapa")<>"" then chapa=request.form("chapa") else chapa=rs("chapa")
%>
<input type="hidden" name="chapa1ant" value="<%=request.form("chapa")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=chapa%>" name="chapa" onchange="chapa1()" size="8" onfocus="javascript:window.status='Informe o chapa do professor'"></td>
	<td class=fundo>
		<select size="1" name="nome" onchange="nome1()" onfocus="javascript:window.status='Selecione o Nome do Funcionario'" >
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' order by nome"
'if session("dp_chapa")<>"" then sql2=sql2 & "and chapa='" & session("dp_chapa") & "'" else sql2=sql2 & "order by nome"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o funcionário....</option>"
rsc.movefirst:do while not rsc.eof
if chapa=rsc("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rsc("chapa")%>" <%=temp%>><%=rsc("nome")%></option>
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
conexao.close
set conexao=nothing

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
%>
</body>
</html>