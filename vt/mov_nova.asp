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
<title>Inclusão de Lançamento VT</title>
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
dim conexao, conexao2, chapach, rs, rs2, ok
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	total=cdbl(request.form("tarifa"))*cint(request.form("quantidade"))

	sqla = "INSERT INTO vt_saldo (data, id_tipo, codigo, tarifa, quantidade, total, chapa "
	sqla = sqla & " )"
	sqlb = " SELECT "
	if request.form("data")="" then sqlb=sqlb & "null" else sqlb=sqlb & "'" & dtaccess(request.form("data")) & "'"
	sqlb=sqlb & "," & request.form("id_tipo")
	sqlb=sqlb & ",'" & request.form("codigop") & "'"
	sqlb=sqlb & "," & nraccess(request.form("tarifa"))
	sqlb=sqlb & "," & nraccess(request.form("quantidade"))
	sqlb=sqlb & "," & nraccess(total)
	sqlb=sqlb & ",'" & request.form("chapa") & "'"
	'sqlb=sqlb & ",'" & session("usuariomaster") & "'"
	'sqlb=sqlb & ",getdate()"
	sql = sqla & sqlb
	'response.write sql
	if tudook=1 then conexao.Execute sql, , adCmdText
end if 'request btsalvar

'if request.form("bt_salvar")="" then
%>
<form method="POST" action="mov_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
	<tr><td class=grupo>Inclusão de Lançamento VT</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
%>
<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Tipo Movimento</td>
	<td class=titulo>Vale Transporte</td>
</tr>
<tr>
<%if request.form("data")<>"" then dataf=request.form("data") else dataf=now%>
	<td class=titulo><input type="text" name="data" size="12" value="<%=formatdatetime(dataf,2)%>"></td>
	<td class=titulo><select size="1" name="id_tipo" onfocus="javascript:window.status='Selecione o tipo de movimento'">
<%
if request.form("id_tipo")<>"" then id_tipo=request.form("id_tipo") else id_tipo=-1
sqla="SELECT * from vt_saldo_tipo where id_tipo<>0"
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
	
	<td class=titulo><select size="1" name="codigop" onChange="javascript:submit()" onfocus="javascript:window.status='Selecione o Vale Transporte'">
<%
if request.form("codigop")<>"" then codigop=request.form("codigop") else codigop=""
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
sql="select valor from ptarifa where codigo='" & codigop & "' "
sql="select valor from corporerm.dbo.ptarifa where codigo='" & codigop & "' and '" & dtaccess(request.form("data")) & "' between iniciovigencia and finalvigencia "
rsd.Open sql, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then tarifa=rsd("valor") else tarifa=0
rsd.close
if request.form("quantidade")<>"" then quantidade=request.form("quantidade") else quantidade=0
if request.form("total")<>"" then total=request.form("total") else total=0
total=cdbl(tarifa)*cint(quantidade)
%>
<input type="hidden" name="totalt" value="<%=total%>">
	<td class=titulo><input type="text" name="tarifa"     class=vr size="8" value="<%=formatnumber(tarifa,2)%>"></td>
	<td class=titulo><input type="text" name="quantidade" class=vr size="8" value="<%=formatnumber(quantidade,0)%>"></td>
	<td class=titulo><input type="text" name="total"      class=vr size="8" value="<%=formatnumber(total,2)%>" onFocus="total.blur()" disabled></td>
</tr>
</table>

<!-- Chapa / Nome -->
<%
if request.form("chapa")<>"" then chapa=request.form("chapa") else chapa=""
%>
<input type="hidden" name="chapa1ant" value="<%=request.form("chapa")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=titulo><input type="text" value="<%=chapa%>" name="chapa" onchange="chapa1()" size="8" onfocus="javascript:window.status='Informe o chapa do professor'"></td>
	<td class=titulo>
		<select size="1" name="nome" onchange="nome1()" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" >
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
'end if   'request.form=""
conexao.close
set conexao=nothing

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
%>
</body>
</html>