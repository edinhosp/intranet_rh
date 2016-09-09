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
<title>Inclusão de Estacionamento</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	if request.form("vy")="ON" then vy=-1 else vy=0
	if request.form("bp")="ON" then bp=-1 else bp=0
	if request.form("ns")="ON" then ns=-1 else ns=0
	if request.form("jw")="ON" then jw=-1 else jw=0
	if request.form("pavy")="-1" then pavy=-1 else pavy=0
	if request.form("pabp")="-1" then pabp=-1 else pabp=0
	if request.form("pans")="-1" then pans=-1 else pans=0
	if request.form("pajw")="-1" then pajw=-1 else pajw=0
	if request.form("paid")<>"" then
		if cdate(request.form("pafim"))>cdate(request.form("inicio")) then
			sql="update veiculos_a set termino='" & dtaccess(cdate(request.form("inicio"))-1) & "' where id_est=" & request.form("paid")
			'response.write sql
			if tudook=1 then conexao.Execute sql, , adCmdText
		end if
	end if
	
	sql = "INSERT INTO veiculos_a (" 
	sql = sql & "chapa, inicio, termino, cartao, obs, vy, bp, ns, jw, pavy, pabp, pans, pajw, usuarioa, dataa "
	sql = sql & ") "
	sql2 = " SELECT "
	sql2=sql2 & " '" & request.form("chapa") & "', "
	if request.form("inicio")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("inicio")) & "', "
	if request.form("termino")=""  then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("termino")) & "', "
	sql2=sql2 & " '" & request.form("cartao") & "', "
	sql2=sql2 & " '" & request.form("obs") & "', "
	sql2=sql2 & " " & vy & ", "
	sql2=sql2 & " " & bp & ", "
	sql2=sql2 & " " & ns & ", "
	sql2=sql2 & " " & jw & ", "
	sql2=sql2 & " " & pavy & ", "
	sql2=sql2 & " " & pabp & ", "
	sql2=sql2 & " " & pans & ", "
	sql2=sql2 & " " & pajw & ", "
	sql2=sql2 & " '" & session("usuariomaster") & "', "
	sql2=sql2 & " getdate() "
	sql1 = sql & sql2 & ""
	'response.write "<font size='2'>" & sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
end if

%>
<form method="POST" action="estac_nova.asp" name="form" >
<input type="hidden" name="pavy" value="<%=request("pavy")%>">
<input type="hidden" name="pans" value="<%=request("pans")%>">
<input type="hidden" name="pabp" value="<%=request("pabp")%>">
<input type="hidden" name="pajw" value="<%=request("pajw")%>">
<input type="hidden" name="paid" value="<%=request("paid")%>">
<input type="hidden" name="pafim" value="<%=request("pafim")%>">
<input type="hidden" name="acartao" value="<%=request("acartao")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Estacionamento</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário</td></tr>
<tr><td class=titulo><select size="1" name="chapa" class=a>
<%
sql2="select chapa, nome from (select chapa, nome, codsecao as descricao, codsindicato from grades_novos union all select f.chapa collate database_default, f.nome collate database_default, f.secao collate database_default, f.codsindicato collate database_default from qry_funcionarios f) as t where chapa='" & request("chapa") & "' " 
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if request("chapa")=rs2("chapa") then tempc="selected" else tempc=""
%>
          <option value="<%=rs2("chapa")%>" <%=tempc%>><%=rs2("nome")%></option>
<%
rs2.movenext
loop
rs2.close
%>
	</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Inicio</td>
	<td class=titulo>Término</td>
	<td class=titulo>Cartão B.P.</td>
</tr>
<%
mesinclusao=month(now)
select case mesinclusao
	case 3,4,5,6,7,8,9,10,11,12
		termino=dateserial(year(now)+1,2,28)
	case 1,2
		termino=dateserial(year(now),2,28)
end select
%>
<tr>
	<td class=titulo><input type="text" name="inicio" size="8" value="<%=formatdatetime(now,2)%>"></td>
	<td class=titulo><input type="text" name="termino"  size="8" value="<%=termino%>" > </td> <!-- onfocus="this.blur()" -->
	<td class=titulo><input type="text" name="cartao" size="5" value="<%=request("acartao")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Coral</td>
	<td class=titulo>Branco</td>
	<td class=titulo>Narciso</td>
	<td class=titulo>J.Wilson</td>
	<td class=titulo>Observações</td>
</tr>
<tr>
	<td class=titulo><input type="checkbox" name="vy" value="ON"></td>
	<td class=titulo><input type="checkbox" name="bp" value="ON"></td>
	<td class=titulo><input type="checkbox" name="ns" value="ON"></td>
	<td class=titulo><input type="checkbox" name="jw" value="ON"></td>
	<td class=titulo><input type="text" name="obs" size="50" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
	<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
	<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
	<input type="button" value="Fechar" class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
'rs.close
set rs=nothing
set rs2=nothing
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