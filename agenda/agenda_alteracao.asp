<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a96")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Compromisso</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1() { form.nome.value=form.nome.value.toUpperCase() }
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

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		if request.form("hora")="" then hora="NULL" else hora="'" & request.form("hora") & "'"
		usuarioc=session("usuariomaster")
		datac=dtaccess(now())
		
		sql="UPDATE agenda SET " & _
		"data   ='" & dtaccess(request.form("data")) & "' " & _
		",hora  = " & hora & " " & _
		",compromisso='" & request.form("anotacao") & "' " & _
		",tipo  = " & request.form("tipo") & " " & _
		",usuarioa='" & session("usuariomaster") & "' " & _
		",dataa   =getdate() " & _
		" WHERE id_agenda=" & session("id_alt_agenda")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("tipo")="1" and tudook=1 then
		idagenda=session("id_alt_agenda")
		sqli="delete from agenda_1 where id_agenda=" & idagenda
		conexao.Execute sqli, , adCmdText
		sqli="delete from agenda_3 where id_agenda=" & idagenda
		conexao.Execute sqli, , adCmdText
		sqli="insert into agenda_1 (id_agenda, codigo) select " & idagenda & ",'" & request.form("codigo1") & "' "
		conexao.Execute sqli, , adCmdText
	end if
	if request.form("tipo")="3" and tudook=1 then
		chapas=request.form("codigo3").count
		idagenda=session("id_alt_agenda")
		sqli="delete from agenda_3 where id_agenda=" & idagenda
		conexao.Execute sqli, , adCmdText
		sqli="delete from agenda_1 where id_agenda=" & idagenda
		conexao.Execute sqli, , adCmdText
		for a=1 to chapas
			sqli="insert into agenda_3 (id_agenda, codigo) select " & idagenda & ",'" & request.form("codigo3").item(a) & "' "
			conexao.Execute sqli, , adCmdText
		next
	end if
	
	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM agenda WHERE id_agenda=" & session("id_alt_agenda")
		sql1="DELETE FROM agenda_1 WHERE id_agenda=" & session("id_alt_agenda")
		sql3="DELETE FROM agenda_3 WHERE id_agenda=" & session("id_alt_agenda")
		if tudook=1 then conexao.Execute sql, , adCmdText
		if tudook=1 then conexao.Execute sql1, , adCmdText
		if tudook=1 then conexao.Execute sql3, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_agenda=session("id_alt_agenda")
		id_agenda=request.form("id_agenda")
	else
		id_agenda=request("codigo")
	end if
	sql="select * from agenda where id_agenda=" & id_agenda
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
session("id_alt_agenda")=rs("id_agenda")
if request.form("data")<>"" then dataanot=request.form("data") else dataanot=rs("data")
if request.form("hora")<>"" then horaanot=request.form("hora") else horaanot=rs("hora")
if isdate(horaanot)=true then horaanot=formatdatetime(horaanot,4)
if isdate(dataanot)=true then diaanot="<font color=blue>" & weekdayname(weekday(dataanot)) else diaanot=""
if request.form("anotacao")<>"" then anotacao=request.form("anotacao") else anotacao=rs("compromisso")
if request.form("tipo")<>"" then tipo=request.form("tipo") else tipo=rs("tipo")
%>
<form method="POST" action="agenda_alteracao.asp" name="form">
<input type="hidden" name="id_agenda" size="4" value="<%=rs("id_agenda")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
	<tr><td class=grupo>Alteração de Compromisso/Lembrete <%=rs("id_agenda")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
<tr>
	<td class=titulo>Data Anotação</td>
	<td class=titulo>Hora (opcional)</td>
</tr>
<tr>
	<td class=fundo><input type="text" size=8 name="data" value="<%=dataanot%>" onChange="javascript:submit();">&nbsp;<%=diaanot%></td>
	<td class=fundo><input type="text" size=4 name="hora" value="<%=horaanot%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
<tr>
	<td class=titulo>Anotação</td>
</tr>
<tr>
	<td class=fundo><textarea name="anotacao" cols="50" rows="3"><%=anotacao%></textarea></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
<tr>
	<td class=titulo>Tipo Anotação</td>
	<td class=titulo>---</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="tipo" onfocus="javascript:window.status='Selecione o tipo de anotação'" onChange="javascript:submit();">
		<option value="0" <%if tipo="0" then response.write "selected"%>>Anotação Pessoal</option>
		<option value="1" <%if tipo="1" then response.write "selected"%>>Para o Setor: </option>
		<%if session("usuariomaster")="02379" then%>
		<option value="2" <%if tipo="2" then response.write "selected"%>>Todos usuários do RH Online</option>
		<%end if%>
		<option value="3" <%if tipo="3" then response.write "selected"%>>Para os usuários: </option>
	</select></td>
	<td class=fundo>
<%
if tipo="1" then 'para grupo especifico
sql1="select grupo from usuarios where usuario='" & session("usuariomaster") & "';"
rsd.Open sql1, ,adOpenStatic, adLockReadOnly
codigo1=rsd("grupo")
rsd.close
%>
	<select size=1 name="codigo1">
		<option value="<%=codigo1%>"><%=codigo1%></option>
<%
sql1="SELECT grupo FROM usuarios WHERE ativo<>0 GROUP BY grupo HAVING grupo Is Not Null and grupo<>'" & codigo1 & "'; "
rsd.Open sql1, ,adOpenStatic, adLockReadOnly
do while not rsd.eof
%>
	<option value="<%=rsd("grupo")%>"><%=rsd("grupo")%></option>
<%
rsd.movenext
loop
rsd.close
%>
	</select>
<%
end if
if tipo="3" then 'para usuario especifico
%>
	<select size=7 name="codigo3" multiple>
<%
sql1="SELECT usuario, nome FROM usuarios WHERE ativo<>0 order by nome; "
rsd.Open sql1, ,adOpenStatic, adLockReadOnly
do while not rsd.eof
%>
	<option value="<%=rsd("usuario")%>"><%=rsd("nome")%></option>
<%
rsd.movenext
loop
rsd.close
%>
	</select>
<%
end if
%>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
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