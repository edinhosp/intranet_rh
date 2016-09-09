<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a79")="N" or session("a79")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Convênio de Bolsas</title>
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
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1

		sqla = "INSERT INTO rhconveniados (chapa, id_faculdade, curso, periodo, data, encaminhamento, " & _
		"obs, anoletivo )"
		
		sqlb = " SELECT '" & request.form("chapa") & "' " & _
		"," & request.form("id_faculdade") & " " & _
		",'" & request.form("curso") & "' " & _
		",'" & request.form("periodo") & "' " & _
		",'" & dtaccess(now()) & "' " & _
		",'" & request.form("encaminhamento") & "' " & _
		",'" & request.form("obs") & "' " & _
		",'" & request.form("anoletivo") & "' " 
		'sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		'sqlb=sqlb & ",getdate()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if 'request btsalvar
else 'request.form=""
end if
if request.form("bt_salvar")<>"" then
else
end if	

'if request.form("bt_salvar")="" then
%>
<form method="POST" action="enviados_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
	<tr><td class=grupo>Inclusão de Convênio de Bolsas</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
if request("codigo")<>"" then
	chapa=request("codigo")
elseif request.form("chapa")<>"" then
	chapa=request.form("chapa") 
else
	chapa=""
end if
%>
<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo>0</td>
	<td class=fundo><input type="text" value="<%=chapa%>" name="chapa" size="5" onfocus="javascript:window.status='Informe o chapa do funcionário'" onchange="chapa1()"></td>
	<td class=fundo>
		<select size="1" name="nome" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" onchange="nome1()">
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' order by nome "
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

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Faculdade</td>
</tr>
<tr>
<%if request.form("ano")<>"" then ano=request.form("ano") else ano=year(now)%>
<%if request.form("mes")<>"" then mes=request.form("mes") else mes=month(now)%>
	<td class=fundo>
		<select size="1" name="id_faculdade" onchange="javascript:submit()">
		<option value="0">Selecione uma faculdade</option>
<%
sql2="select id_faculdade, faculdade from rhconveniobe order by faculdade "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if cint(request.form("id_faculdade"))=cint(rsc("id_faculdade")) then temp1="selected" else temp1=""
%>
		<option value="<%=rsc("id_faculdade")%>" <%=temp1%>><%=rsc("faculdade")%></option>
<%
rsc.movenext
loop
rsc.close
%>
	</select>
	</td>
</tr>	
</table>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Curso</td>
</tr>
<tr>
	<td class=fundo>
	<select size="1" name="curso">
		<option value="0">Selecione um curso</option>
        <%
if request.form("id_faculdade")="" then id_faculdade=0 else id_faculdade=request.form("id_faculdade")		
sql2="select cursos, id_curso from rhconveniobec where id_faculdade=" & id_faculdade & " order by cursos "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then 
rsc.movefirst:do while not rsc.eof
if request.form("curso")=rsc("cursos") then temp1="selected" else temp1=""
%>
		<option value="<%=rsc("cursos")%>" <%=temp1%>><%=rsc("cursos")%></option>
<%
rsc.movenext:loop
end if 'recordcount
rsc.close
%>
	</select>	
	</td>
</tr>
</table>

<!-- Periodo/Tipo -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Período</td>
	<td class=titulo>Tipo</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="periodo">
		<option value="0">Selecione um período</option>
		<option value="Matutino"   <%if request.form("periodo")="Matutino"   then response.write "selected"%>>Matutino</option>
		<option value="Vespertino" <%if request.form("periodo")="Vespertino" then response.write "selected"%>>Vespertino</option>
		<option value="Noturno"    <%if request.form("periodo")="Noturno"    then response.write "selected"%>>Noturno</option>
	</select>	
	</td>
	<td class=fundo><select size="1" name="encaminhamento">
		<option value="0">Selecione um tipo</option>
		<option value="1" <%if request.form("encaminhamento")="1" then response.write "selected"%>>Inscrição no Vestibular</option>
		<option value="2" <%if request.form("encaminhamento")="2" then response.write "selected"%>>Matrícula</option>
		<option value="3" <%if request.form("encaminhamento")="3" then response.write "selected"%>>Renovação de Matrícula</option>
	</select>	
	</td>
</tr>
</table>

<!-- Periodo/Tipo -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Ano Letivo</td>
	<td class=titulo>Observações</td>
</tr>
<tr>
	<td class=fundo valign="top"><input type="text" size="6" name="anoletivo" value="<%=request.form("anoletivo")%>">
	</td>
	<td class=fundo><textarea name="obs" cols="30" rows="2"><%=request.form("obs")%></textarea>
	</td>
</tr>
</table>


<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
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