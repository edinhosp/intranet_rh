<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a64")="N" or session("a64")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Convênio de Bolsas</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
function evento1()	{	form.codevento.value=form.evento.value;	form.submit();	}
function codigo1()	{	form.evento.value=form.codevento.value;	form.submit();	}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpro(2), varper(2)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		sql="UPDATE bolsistas_lanc SET " & _
		"renovacao      ='" & dtaccess(request.form("renovacao")) & "' " & _
		",situacao      ='" & request.form("situacao") & "' " & _
		",ano_letivo    ='" & request.form("ano_letivo") & "' " & _ 
		",observacao    ='" & request.form("observacao") & "' " & _
		",protocolo     = " & request.form("protocolo") & " " & _
		",id_faculdade  = " & request.form("id_faculdade") & " " & _
		",periodo       ='" & request.form("periodo") & "' " & _ 
		",curso         ='" & request.form("curso") & "' "
		if request.form("validade")="" then sql=sql & ",validade=null" else sql=sql & "validade='" & dtaccess(request.form("validade")) & "'"
		'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		'sql=sql & ",dataa   =now() "
		sql=sql & " WHERE id_lanc=" & session("id_alt_lanc")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		'response.write sql
		if tudook=1 then 
			conexao.Execute sql, , adCmdText
			id_bolsa1=request.form("id_bolsa")
			sql2="select faculdade from rhconveniobe where id_faculdade=" & request.form("id_faculdade")
			rs.Open sql2, ,adOpenStatic, adLockReadOnly
			if rs.recordcount>0 then faculdade=rs("faculdade") else faculdade="-"
			rs.close
		sql1="select top 1 renovacao from bolsistas_lanc where id_bolsa=" & id_bolsa1 & " order by renovacao desc"
		rsc.Open sql1, ,adOpenStatic, adLockReadOnly
		if rsc.recordcount>0 then
			ultimolanc=rsc("renovacao")
			if cdate(ultimolanc)=cdate(request.form("renovacao")) then
			sql2="update bolsistas set situacao='" & request.form("situacao") & "', instituicao='" & faculdade & "', curso='" & request.form("curso") & "' where id_bolsa=" & id_bolsa1
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
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")=null or request("codigo")="" then
		id_lanc=session("id_alt_lanc")
		if session("id_alt_lanc")="" then id_lanc=request.form("id_lanc")
	else
		id_lanc=request("codigo")
	end if
	sql="select * from bolsistas_lanc where id_lanc=" & id_lanc
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_lanc")=rs("id_lanc")
'sqlz="select nome from pfunc where chapa='" & rs("chapa") & "'"
'set rsnome=server.createobject ("ADODB.Recordset")
'set rsnome=conexao.Execute (sqlz, , adCmdText)
'response.write request.form
if request.form("situacao")=""    then situacao=rs("situacao")       else situacao=request.form("situacao")
if request.form("ano_letivo")=""  then ano_letivo=rs("ano_letivo")   else ano_letivo=request.form("ano_letivo")
if request.form("renovacao")=""   then renovacao=rs("renovacao")     else renovacao=request.form("renovacao")
if request.form("validade")=""    then validade=rs("validade")       else validade=request.form("validade")
if request.form("observacao")=""  then observacao=rs("observacao")   else observacao=request.form("observacao")
if request.form("protocolo")=""   then protocolo=rs("protocolo")     else protocolo=request.form("protocolo")
if request.form("id_faculdade")="" then faculdade=rs("id_faculdade") else faculdade=request.form("id_faculdade")
if request.form("curso")=""       then curso=rs("curso")             else curso=request.form("curso")
if request.form("periodo")=""     then periodo=rs("periodo")         else periodo=request.form("periodo")

%>
<form method="POST" action="convenio_alteracao.asp" name="form">
<input type="hidden" name="id_lanc" size="4" value="<%=rs("id_lanc")%>">
<input type="hidden" name="id_bolsa" size="4" value="<%=rs("id_bolsa")%>" >  
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
	<tr><td class=grupo>Alteração de Convênio de Bolsa <%=rs("id_lanc")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Data Renovação</td>
	<td class=titulo>Data Validade</td>
	<td class=titulo>Situação do Lançamento</td>
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
</tr>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Periodo Letivo</td>
	<td class=titulo>Observação</td>
	<td class=titulo>Tipo Emissão</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="ano_letivo"  size="15" value="<%=ano_letivo%>"></td>
	<td class=fundo><input type="text" name="observacao"  size="30" value="<%=observacao%>"></td>
	<td class=fundo><select size="1" name="protocolo">
          <option value="-1">Selecione</option>
<%
varpro(0)="Inscr.Vestibular"
varpro(1)="Matrícula"
varpro(2)="Rematrícula"
for a=0 to 2
	if cdbl(protocolo)=cdbl(a+1) then tempp="selected" else tempp=""
%>
	<option value="<%=a+1%>" <%=tempp%>><%=varpro(a)%></option>
<%
next
%>
	</select></td>
</tr>
</table>
<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Faculdade</td>
	<td class=titulo>Período</td>
</tr>
<tr>
	<td class=fundo>
		<select size="1" name="id_faculdade" onchange="javascript:submit()">
		<option value="0">Selecione uma faculdade</option>
<%
sql2="select id_faculdade, faculdade from rhconveniobe order by faculdade "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if cint(faculdade)=cint(rsc("id_faculdade")) then temp1="selected" else temp1=""
%>
		<option value="<%=rsc("id_faculdade")%>" <%=temp1%>><%=rsc("faculdade")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select></td>
	<td class=fundo><select size="1" name="periodo">
          <option value=""></option>
<%
varper(0)="Matutino"
varper(1)="Vespertino"
varper(2)="Noturno"
for a=0 to 2
	if periodo=varper(a) then tempp="selected" else tempp=""
%>
	<option value="<%=varper(a)%>" <%=tempp%>><%=varper(a)%></option>
<%
next
%>
	</select></td>
</tr>	
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Curso</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="curso">
		<option value="0">Selecione um curso</option>
<%
sql2="select cursos, id_curso from rhconveniobec where id_faculdade=" & faculdade & " order by cursos "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then 
rsc.movefirst:do while not rsc.eof
if curso=rsc("cursos") then temp1="selected" else temp1=""
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

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
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


set rsc=nothing
set rsd=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>