<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a61")="" or session("a61")="N" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Alteração de Controle de Teto</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

	if request.form("bt_salvar")<>"" then
		tudook=1
		sql="UPDATE rhcontroleteto SET "
		sql=sql & "ano          = " & request.form("ano")      & ", "
		sql=sql & "mes          = " & request.form("mes")      & ", "
		sql=sql & "empresa      = '" & request.form("empresa") & "', "
		sql=sql & "cnpj         = '" & request.form("cnpj")    & "', "
		sql=sql & "proporcional = "  & nraccess(request.form("proporcional")) & ", "
		if request.form("data")<>"" then sql=sql & "data='" & dtaccess(request.form("data")) & "', " else sql=sql & "data=null, "
		if request.form("copia")="ON" then sql=sql & "copia=1, " else sql=sql & "copia=0, "
		sql=sql & "usuarioa='" & session("usuariomaster") & "', "
		sql=sql & "dataa=getdate() "
		sql=sql & "WHERE id_teto=" & request.form("id_teto")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM rhcontroleteto WHERE id_teto=" & request.form("id_teto")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null then
		id_teto=session("id_alt_teto")
	else
		id_teto=request("codigo")
	end if
	sql1="select * from rhcontroleteto where id_teto=" & id_teto
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_teto")=rs("id_teto")
sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
%>
<form method="POST" action="teto_alteracao.asp" name="form">
<input type="hidden" name="id_teto" size="4" value="<%=rs("id_teto")%>" style="font-size: 8 pt" >  
<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr><td class=grupo>Alteração de Controle de Teto Máximo</td></tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=titulo><%=rs("id_teto")%></td>
	<td class=titulo><%=rs("chapa")%></td>
	<td class=titulo><%=rsnome("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr>
	<td class=titulo>Ano</td>
	<td class=titulo>Mês</td>
	<td class=titulo>Data Entrega</td>
	<td class=titulo>é Cópia/Fax?</td>
</tr>
<tr>
	<td class=titulo><select size="1" name="ano">
	<%for ano=year(now)-1 to year(now)+1%>
		<option value="<%=ano%>" <%if ano=cint(rs("ano")) then response.write "Selected"%>><%=ano%></option>
	<%next%>
		</select>
	</td>
	<td class=titulo><select size="1" name="mes">
	<%for mes=1 to 13%>
		<option value="<%=mes%>" <%if mes=rs("mes") then response.write "Selected"%>><%=mes%></option>
	<%next%>
		</select>
	</td>
	<td class=titulo><input type="text" name="data" size="12" value="<%=rs("data")%>" ></td>
<%if rs("copia")=-1 then copia="checked" else copia=""%>
	<td class=titulo><input type="checkbox" name="copia" value="ON" <%=copia%> ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr>
	<td class=titulo>Empresa</td>
	<td class=titulo>CNPJ</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="empresa"  size="35" value="<%=rs("empresa")%>"></td>
	<td class=titulo><input type="text" name="cnpj"  size="20" value="<%=rs("cnpj")%>"></td>
</tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr>
	<td class=titulo>Contribuição Proporcional</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="proporcional" size="15" value="<%=formatnumber(rs("proporcional"),2) %>" ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
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