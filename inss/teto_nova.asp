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
<title>Inclusão de Controle de Teto</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
--></script>
</head>
<body>
<script src="../coolmenu/coolmenus_frame.js" type="text/javascript"></script>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
if request.form("bt_salvar")<>"" then
	tudook=1
	if request.form("chapa")="" or request.form("ano")="" or request.form("mes")=""  then _
		tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Preencha os campos da chapa, ano ou mês!!');</script>"
	
	sql = "INSERT INTO rhcontroleteto ("
	sql = sql & "chapa, ano, mes, empresa, cnpj, data, proporcional, copia, usuarioc, datac"
	sql = sql & ") "
	if request.form("proporcional")="" then proporcional=0  else proporcional=request.form("proporcional")
	sql2 = " SELECT '" & request.form("chapa") & "', "
	sql2=sql2 & " " & request.form("ano") & ", "
	sql2=sql2 & " " & request.form("mes") & ", "
	sql2=sql2 & " '" & request.form("empresa") & "', "
	sql2=sql2 & " '" & request.form("cnpj") & "', "
	if request.form("data")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("data")) & "', "
	sql2=sql2 & " " & nraccess(proporcional) & ", "
	if request.form("copia")="ON" then sql2=sql2 & "-1, " else sql2=sql2 & "0, "
	sql2=sql2 & " '" & session("usuariomaster") & "', "
	sql2=sql2 & " getdate() "
	sql4 = sql & sql2 & ""
	'response.write "<font size='2'>" & sql4
	if tudook=1 then conexao.Execute sql4, , adCmdText
end if

else 'request.form=""
end if

'if request.form="" then
if request("codigo")<>"" then chapainc=request("codigo")
%>
<form method="POST" action="teto_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr><td class=grupo>Inclusão de Controle de Teto Máximo</td></tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=titulo>0</td>
	<td class=titulo><input type="text" name="chapa" size=5 value="<%=chapainc%>" onchange="chapa1()"></td>
	<td class=titulo><select size="1" name="nome" onchange="nome1()">
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and codtipo='N' order by nome"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if rsc("chapa")=chapainc then temp="selected" else temp=""
%>
		<option value="<%=rsc("chapa")%>" <%=temp%>><%=rsc("nome")%></option>
<%
rsc.movenext
loop
rsc.close
set rsc=nothing
%>
	</select></td>
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
		<option value="<%=ano%>" <%if ano=year(now) then response.write "Selected"%>><%=ano%></option>
	<%next%>
		</select>
	</td>
	<td class=titulo><select size="1" name="mes">
	<%for mes=1 to 13%>
		<option value="<%=mes%>" <%if mes=month(now) then response.write "Selected"%>><%=mes%></option>
	<%next%>
		</select>
	</td>
	<td class=titulo><input type="text" name="data" size="12" value="<%=formatdatetime(now,2)%>"></td>
	<td class=titulo><input type="checkbox" name="copia" value="ON"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr>
	<td class=titulo>Empresa</td>
	<td class=titulo>CNPJ</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="empresa" size="40" value="<%=session("tetoultimaempresa")%>"></td>
	<td class=titulo><input type="text" name="cnpj" size="20" value="<%=session("tetoultimocnpj")%>"></td>
	</tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr>
	<td class=titulo>Contribuição Proporcional</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="proporcional" size="15" ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="430">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="button" value="Fechar" class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing
'end if
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