<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a81")="" or session("a81")="N" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Assistência Médica</title>
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
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	sql="UPDATE assmed_mudanca SET "
	sql=sql & "empresa  = '"   & request.form("empresa") & "', "
	sql=sql & "plano    = '"   & request.form("plano") & "', "
	sql=sql & "codigo   = '"   & request.form("codigo1") & "', "
	sql=sql & "up       = '"   & request.form("up") & "', "
	if request.form("inclusao")<>"" then 
		sql=sql & "inclusao = '"   & dtaccess(request.form("inclusao")) & "', "
	else
		sql=sql & "inclusao = null, "
	end if
	if request.form("ivigencia")<>"" then 
		sql=sql & "ivigencia = '"   & dtaccess(request.form("ivigencia")) & "', "
	else
		sql=sql & "ivigencia = null, "
	end if
	if request.form("fvigencia")<>"" then 
		sql=sql & "fvigencia = '"   & dtaccess(request.form("fvigencia")) & "', "
	else
		sql=sql & "fvigencia = '12/31/2020', "
	end if
	if request.form("compr")="ON" then 
		sql=sql & "compr = 1, " 
	else
		sql=sql & "compr = 0, "
	end if
	sql=sql & "oper      = '"   & request.form("oper")   & "', "
	sql=sql & "uoper     = '"   & request.form("uoper")  & "' "
	sql=sql & " WHERE id_mudanca=" & session("id_alt_mudanca")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM assmed_mudanca WHERE id_mudanca=" & session("id_alt_mudanca")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_mudanca=session("id_alt_mudanca")
	else
		id_mudanca=request("codigo")
	end if
	sqla="select * from assmed_mudanca "
	sqlb="where id_mudanca=" & id_mudanca
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_mudanca")=rs("id_mudanca")

sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
%>
<form method="POST" action="controle_alteracao.asp" name="form">
<input type="hidden" name="id_mudanca" size="4" value="<%=rs("id_mudanca")%>" style="font-size: 8 pt" >  
<table border="0" cellpadding="1" cellspacing="1" width="500">
<tr>
	<td class=grupo>Assistência Médica</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo>Funcionário Titular</td></tr>
<tr><td class=titulo><p class=realce><%=rs("chapa")%> - <%=rsnome("nome")%></p></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Empresa de saúde</td>
	<td class=titulo>Plano escolhido</td>
	<td class=titulo>UP</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="empresa" onchange="javascript:submit()">
	<option value="">Selecione...</option>
<%
if request.form("empresa")="" then empresa=rs("empresa") else empresa=request.form("empresa")
sqla="SELECT * from assmed_empresa where ativo=1 ORDER by operadora"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:	do while not rsc.eof
if rsc("codigo")=empresa then tempt="selected" else tempt=""
%>
	<option value="<%=rsc("codigo")%>" <%=tempt%>><%=rsc("operadora")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select></td>
	<td class=fundo>
		<select size="1" name="plano">
			<option value="">Selecione um plano de saúde</option>
<%
if request.form("plano")="" then plano=rs("plano") else plano=request.form("plano")
if request.form("up")=""    then up=rs("up")       else up=request.form("up")
sqla="SELECT * from assmed_planos where codigo='" & empresa & "' ORDER by seq, plano"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rsc("plano")=plano then tempp="selected" else tempp=""
%>
		<option value="<%=rsc("plano")%>" <%=tempp%>><%=rsc("plano") %></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select>
	</td>
	<td class=titulo><input type="text" name="up" size="4" maxlenght=4 value="<%=up%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Código da carteirinha</td>
	<td class=titulo>Inclusão</td>
	<td class=titulo>Início Cobr.</td>
	<td class=titulo>Término</td>
</tr>
<tr>
<%
if request.form("codigo1")=""    then codigo1  =rs("codigo")    else codigo1   =request.form("codigo1")
if request.form("inclusao")="" then inclusao=rs("inclusao") else inclusao=request.form("inclusao")
if request.form("ivigencia")="" then ivigencia=rs("ivigencia") else ivigencia=request.form("ivigencia")
if request.form("fvigencia")="" then fvigencia=rs("fvigencia") else fvigencia=request.form("fvigencia")
if len(codigo1)=16 then tamanho="" else tamanho="Cod.Inc."
%>
	<td class=titulo><input type="text" name="codigo1" size="16" maxlenght=16 value="<%=codigo1%>"><%=tamanho%></td>
	<td class=titulo><input type="text" name="inclusao" size="12" value="<%=inclusao%>"></td>
	<td class=titulo><input type="text" name="ivigencia" size="12" value="<%=ivigencia%>"></td>
	<td class=titulo><input type="text" name="fvigencia" size="12" value="<%=fvigencia%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Operação atual de  Cadastro</td>
	<td class=titulo>Ultima operação</td>
	<td class=titulo>Emitiu Comprovante</td>
</tr>
<tr>
<%
if request.form("oper")=""  then oper =rs("oper")  else oper =request.form("oper")
if request.form("uoper")="" then uoper=rs("uoper") else uoper=request.form("uoper")
if rs("compr")=0 and request.form("compr")<>"ON" then obs1="" else obs1="checked"
%>
	<td class=fundo><select size="1" name="oper">
		<option value=""  <%if oper="" then response.write "selected"%>></option>
		<option value="I" <%if oper="I" then response.write "selected"%>>Inclusão</option>
		<option value="A" <%if oper="A" then response.write "selected"%>>Alteração</option>
		<option value="E" <%if oper="E" then response.write "selected"%>>Exclusão</option>
		<option value="2" <%if oper="2" then response.write "selected"%>>2ª Via</option>
		</select>
	</td>
	<td class=fundo><select size="1" name="uoper">
		<option value=""  <%if uoper="" then response.write "selected"%>></option>
		<option value="I" <%if uoper="I" then response.write "selected"%>>Inclusão</option>
		<option value="A" <%if uoper="A" then response.write "selected"%>>Alteração</option>
		<option value="E" <%if uoper="E" then response.write "selected"%>>Exclusão</option>
		<option value="2" <%if uoper="2" then response.write "selected"%>>2ª Via</option>
		</select>
	</td>
	<td class=titulo><input type="checkbox" name="compr" value="ON" <%=obs1 %>>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
      <td class=titulo align="center">
		<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar">
      </td>
      <td class=titulo align="center">
       <input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
      <td class=titulo align="center">
       <input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
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