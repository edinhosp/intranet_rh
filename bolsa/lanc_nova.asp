<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a64")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Lançamento Bolsa de Estudo</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function chapa1() {	form.chapa.value=form.nome.value;	}
function nome1() {	form.nome.value=form.chapa.value;	}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(4), varcur(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
if request.form("bt_salvar")<>"" then
	tudook=1
	'response.write request.form
	if request.form("protocolo")="ON" then compl = 1 else compl = 0
	sql = "INSERT INTO bolsistas_lanc (" 
	sql = sql & "id_bolsa, situacao, ano_letivo, renovacao, validade, observacao, excecao, protocolo "
	sql = sql & ") "
	sql2 = " SELECT "
	sql2=sql2 & " " & request.form("id_bolsa") & ", "
	sql2=sql2 & " '" & request.form("situacao") & "', "
	sql2=sql2 & " '" & request.form("ano_letivo") & "', "
	if request.form("renovacao")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("renovacao")) & "', "
	if request.form("validade")<>"" then validade="'" & dtaccess(request.form("validade")) & "',"
	if request.form("validade")="" then
		mes1=month(request.form("renovacao")):ano1=year(request.form("renovacao"))
		if mes1>=1 and mes1<=5 then mes2=6:data2=dateserial(ano1,mes2,30)
		if mes1>=6 and mes1<=12 then mes2=12:data2=dateserial(ano1,mes2,31)
		validade="'" & dtaccess(data2) & "', "
	end if
	sql2=sql2 & validade
	sql2=sql2 & " '" & request.form("observacao") & "', "
	sql2=sql2 & " '" & request.form("excecao") & "' "
	sql2=sql2 & ", " & compl & " "
	sql1 = sql & sql2 & ""
	'response.write "<font size='1'>" & sql1
	if tudook=1 then 
		conexao.Execute sql1, , adCmdText
		id_bolsa1=request.form("id_bolsa")
		sql1="select top 1 renovacao from bolsistas_lanc where id_bolsa=" & id_bolsa1 & " order by renovacao desc"
		rsc.Open sql1, ,adOpenStatic, adLockReadOnly
		if rsc.recordcount>0 then
			ultimolanc=rsc("renovacao")
			if cdate(ultimolanc)=cdate(request.form("renovacao")) then
				sql2="update bolsistas set situacao='" & request.form("situacao") & "' where id_bolsa=" & id_bolsa1
				conexao.Execute sql2, , adCmdText
			end if
		end if
		rsc.close
	end if
end if
else 'request.form=""
end if

if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then

situacao=request.form("situacao")
ano_letivo=request.form("ano_letivo")
renovacao=request.form("renovacao")
validade=request.form("validade")
observacao=request.form("observacao")
excecao=request.form("excecao")
obs2=request.form("protocolo")
if obs2="ON" then obs1="checked" else obs1=""

if request("codigo")="" or isnull(request("codigo")) then
	codigo=request.form("id_bolsa")
else
	codigo=request("codigo")
end if
%>
<form method="POST" action="lanc_nova.asp" name="form" >
<input type="hidden" name="id_bolsa" size="4" value="<%=codigo%>" >  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Lançamento Bolsa de Estudo <%=codigo%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Data Renovação</td>
	<td class=titulo>Data Validade</td>
	<td class=titulo>Situação do Lançamento</td>
	<td class=titulo>Protocolo Emitido</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="renovacao" value="<%=renovacao%>" size="10"></td>
	<td class=fundo><input type="text" name="validade" value="<%=validade%>" size="10"></td>
	<td class=fundo><select size="1" name="situacao">
<%
sqla="SELECT * from bolsistas_situacao ORDER by descricao"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rsc("id_sit")=situacao then temps="selected" else temps=""
%>
		<option value="<%=rsc("id_sit")%>" <%=temps%>><%=rsc("descricao")%></option>
<%
rsc.movenext
loop
rsc.close
%>
	</select></td>
	<td class=fundo><input type="checkbox" name="protocolo" value="ON" <%=obs1 %>></td>
</tr>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Periodo Letivo</td>
	<td class=titulo>Exceção</td>
	<td class=titulo>Observação</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="ano_letivo"  size="10" value="<%=ano_letivo%>"></td>
	<td class=fundo><input type="text" name="excecao"  size="30" value="<%=excecao%>"></td>
	<td class=fundo><input type="text" name="observacao"  size="30" value="<%=observacao%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
else
'rs.close
set rs=nothing
end if
set rsc=nothing
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