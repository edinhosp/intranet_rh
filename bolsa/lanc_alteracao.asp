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
<title>Alteração de Lançamento Bolsa de Estudo</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function renovacao1()	{ form.urenovacao.value=form.renovacao_anterior.value;	}
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

if request.form("bt_salvar")<>"" then
	tudook=1
	sql="UPDATE bolsistas_lanc SET "
	sql=sql & "situacao      = '"   & request.form("situacao")      & "', "
	sql=sql & "ano_letivo    = '"   & request.form("ano_letivo")    & "', "
	if request.form("renovacao")<>"" then 
		sql=sql & "renovacao='" & dtaccess(request.form("renovacao"))  & "', "
	else
		sql=sql & "renovacao=null, "
	end if
	if request.form("validade")<>"" then 
		validade="validade = '" & dtaccess(request.form("validade"))  & "', "
	else
		validade="validade = null, "
	end if
	if request.form("validade")="" then
		mes1=month(request.form("renovacao")):ano1=year(request.form("renovacao"))
		if mes1>=1 and mes1<=5 then mes2=6:data2=dateserial(ano1,mes2,30)
		if mes1>=6 and mes1<=12 then mes2=12:data2=dateserial(ano1,mes2,31)
		validade="validade=#" & dtaccess(data2) & "#, "
	end if
	sql=sql & validade
	if request.form("protocolo")="ON" then 
		sql=sql & "protocolo = 1, " 
	else
		sql=sql & "protocolo = 0, "
	end if
	
	sql=sql & "observacao    = '" & request.form("observacao") & "', "
	sql=sql & "excecao       = '" & request.form("excecao") & "' "
	sql=sql & " WHERE id_lanc=" & session("id_alt_lanc")
	if tudook=1 then 
		conexao.Execute sql, , adCmdText
		id_bolsa1=request.form("id_bolsa")
		response.write id_bolsa1
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

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM bolsistas_lanc WHERE id_lanc=" & session("id_alt_lanc")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null or request("codigo")="" then
		id_lanc=session("id_alt_lanc")
		if session("id_alt_lanc")="" then id_lanc=request.form("id_lanc")
	else
		id_lanc=request("codigo")
	end if
	sqla="select * from bolsistas_lanc "
	sqlb="where id_lanc=" & id_lanc
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_lanc")=rs("id_lanc")
if rs("protocolo")=0 then obs1="" else obs1="checked"
'sqlz="select nome from pfunc where chapa='" & rs("chapa") & "'"
'set rsnome=server.createobject ("ADODB.Recordset")
'set rsnome=conexao.Execute (sqlz, , adCmdText)

if request.form("situacao")=""    then situacao=rs("situacao")       else situacao=request.form("situacao")
if request.form("ano_letivo")=""  then ano_letivo=rs("ano_letivo")   else ano_letivo=request.form("ano_letivo")
if request.form("renovacao")=""   then renovacao=rs("renovacao")     else renovacao=request.form("renovacao")
if request.form("validade")=""    then validade=rs("validade")       else validade=request.form("validade")
if request.form("observacao")=""  then observacao=rs("observacao")   else observacao=request.form("observacao")
if request.form("excecao")=""     then excecao=rs("excecao")         else excecao=request.form("excecao")
if request.form("protocolo")=""   then obs2=rs("protocolo")          else obs2=request.form("protocolo")
if obs2<>0 or obs2="ON" then obs1="checked" else obs1=""

%>
<form method="POST" action="lanc_alteracao.asp" name="form">
<input type="hidden" name="id_lanc" size="4" value="<%=rs("id_lanc")%>" >  
<input type="hidden" name="id_bolsa" size="4" value="<%=rs("id_bolsa")%>" >  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Bolsa de Estudo&nbsp;</td></tr>
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
		<input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if
set rsc=nothing
set rsnome=nothing
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