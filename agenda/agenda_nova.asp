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
<title>Inclusão de Compromisso</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
--></script>

</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, ok
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

'response.write request.form
'response.write "<br>" & request.form("codigo3")
'response.write "<br>" & request.form("codigo3").count

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1

if request.form("data")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe uma data!');</script>"
if request.form("anotacao")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe um compromisso ou anotação!');</script>"
'if request.form("bolsa")="ON" then bolsa=-1 else bolsa=0

		if request.form("hora")="" then hora="NULL" else hora="#" & request.form("hora") & "#"
		usuarioc=session("usuariomaster")
		datac=now()
		sqla = "INSERT INTO agenda (data, hora, compromisso, tipo, usuarioc, datac) "
		sqlb= " SELECT '" & dtaccess(request.form("data")) & "'"
		sqlb=sqlb & ", '" & request.form("hora") & "'"
		sqlb=sqlb & ", '" & request.form("anotacao") & "'"
		sqlb=sqlb & ", " & request.form("tipo")
		sqlb=sqlb & ",'" & usuarioc & "'"
		sqlb=sqlb & ", '" & dtaccess(datac) & "' "
		sql = sqla & sqlb
		if tudook=1 then conexao.Execute sql, , adCmdText
		if request.form("tipo")="1" and tudook=1 then
			sqlr="select id_agenda from agenda where data='" & dtaccess(request.form("data")) & "' and compromisso='" & request.form("anotacao") & "' and usuarioc='" & usuarioc & "' and datac='" & dtaccess(datac) & "' and tipo=" & request.form("tipo") & " order by id_Agenda desc "
			rs.Open sqlr, ,adOpenStatic, adLockReadOnly
			idagenda=rs("id_agenda")
			rs.close
			sqli="insert into agenda_1 (id_agenda, codigo) select " & idagenda & ",'" & request.form("codigo1") & "'; "
			response.write sqli
			conexao.Execute sqli, , adCmdText
		end if
		if request.form("tipo")="3" and tudook=1 then
			chapas=request.form("codigo3").count
			sqlr="select id_agenda from agenda where data='" & dtaccess(request.form("data")) & "' and compromisso='" & request.form("anotacao") & "' and usuarioc='" & usuarioc & "' and datac='" & dtaccess(datac) & "' and tipo=" & request.form("tipo") & " order by id_Agenda desc "
			rs.Open sqlr, ,adOpenStatic, adLockReadOnly
			idagenda=rs("id_agenda")
			rs.close
			for a=1 to chapas
				sqli="insert into agenda_3 (id_agenda, codigo) select " & idagenda & ",'" & request.form("codigo3").item(a) & "' "
				conexao.Execute sqli, , adCmdText
			next
		end if

	end if 'request btsalvar
else 'request.form=""
end if

'if request.form("bt_salvar")="" then
'if request.form("data")="" then dataanot=formatdatetime(now,2) else dataanot=request.form("data")
if request.form("data")="" and request("data")<>"" then dataanot=request("data") else dataanot=request.form("data")
if dataanot="" then dataanot=formatdatetime(now,2)
horaanot=request.form("hora")
if isdate(dataanot)=true then diaanot="<font color=blue>" & weekdayname(weekday(dataanot)) else diaanot=""
anotacao=request.form("anotacao")
tipo=request.form("tipo")
%>

<form method="POST" action="agenda_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
	<tr><td class=grupo>Inclusão de Compromisso/Lembrete</td></tr>
</table>
<!-- tipo lancamento -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="465">
<tr>
	<td class=titulo>Data Anotação</td>
	<td class=titulo>Hora (opcional)</td>
</tr>
<tr>
	<td class=fundo><input type="text" size=8 name="data" value="<%=dataanot%>" onChange="javascript:submit();">&nbsp;<%=diaanot%></td>
	<td class=fundo><input type="text" size=4 name="hora" value="<%=horaaanot%>"></td>
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
rs.Open sql1, ,adOpenStatic, adLockReadOnly
codigo1=rs("grupo")
rs.close
%>
	<select size=1 name="codigo1">
		<option value="<%=codigo1%>"><%=codigo1%></option>
<%
sql1="SELECT grupo FROM usuarios WHERE ativo<>0 GROUP BY grupo HAVING grupo Is Not Null and grupo<>'" & codigo1 & "'; "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
	<option value="<%=rs("grupo")%>"><%=rs("grupo")%></option>
<%
rs.movenext
loop
rs.close
%>
	</select>
<%
end if
if tipo="3" then 'para usuario especifico
%>
	<select size=7 name="codigo3" multiple>
<%
sql1="SELECT usuario, nome FROM usuarios WHERE ativo<>0 order by nome; "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
	<option value="<%=rs("usuario")%>"><%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
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