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
<title>Alteração de Bolsa de Estudo</title>
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
	sql="UPDATE bolsistas SET "
	sql=sql & "tp_bolsa      = '"   & request.form("tp_bolsa")      & "', "
	sql=sql & "nome_bolsista = '"   & request.form("nome_bolsista") & "', "
	sql=sql & "parentesco    = '"   & request.form("parentesco")    & "', "
	if request.form("dtnasc")<>"" then 
		sql=sql & "dtnasc = '"   & dtaccess(request.form("dtnasc"))  & "', "
	else
		sql=sql & "dtnasc = null, "
	end if
	sql=sql & "situacao      = '"   & request.form("situacao")      & "', "
	sql=sql & "tipocurso     = '"   & request.form("tipocurso")     & "', "
	sql=sql & "curso         = '"   & request.form("curso")         & "', "
	sql=sql & "instituicao   = '"   & request.form("instituicao")   & "', "
	sql=sql & "matricula     = '"   & request.form("matricula")     & "', "
	sql=sql & "observacao    = '"   & request.form("observacao")    & "', "
	if request.form("comprovante")="ON" then 
		sql=sql & "comprovante = -1 " 
	else
		if request.form("parentesco")="Titular" then
			sql=sql & "comprovante = -1 " 
		else
			sql=sql & "comprovante = 0 "
		end if
	end if
	sql=sql & " WHERE id_bolsa=" & session("id_alt_bolsa")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM bolsistas WHERE id_bolsa=" & session("id_alt_bolsa")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null or request("codigo")="" then
		id_bolsa=session("id_alt_bolsa")
		if session("id_alt_bolsa")="" then id_bolsa=request.form("id_bolsa")
	else
		id_bolsa=request("codigo")
	end if
	sqla="select * from bolsistas "
	sqlb="where id_bolsa=" & id_bolsa
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_bolsa")=rs("id_bolsa")
if rs("comprovante")=0 then obs1="" else obs1="checked"
sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)

if request.form("parentesco")=""    then parentesco=rs("parentesco")       else parentesco=request.form("parentesco")
if request.form("nome_bolsista")="" then nome_bolsista=rs("nome_bolsista") else nome_bolsista=request.form("nome_bolsista")
if request.form("dtnasc")=""        then dtnasc=rs("dtnasc")               else dtnasc=request.form("dtnasc")
if request.form("situacao")=""      then situacao=rs("situacao")           else situacao=request.form("situacao")
if request.form("curso")=""         then curso=rs("curso")                 else curso=request.form("curso")
if request.form("instituicao")=""   then instituicao=rs("instituicao")     else instituicao=request.form("instituicao")
if request.form("tipocurso")=""     then tipocurso=rs("tipocurso")         else tipocurso=request.form("tipocurso")
if request.form("observacao")=""    then observacao=rs("observacao")       else observacao=request.form("observacao")
if request.form("matricula")=""     then matricula=rs("matricula")         else matricula=request.form("matricula")
if request.form("comprovante")=""   then obs2=rs("comprovante")            else obs2=request.form("comprovante")
if obs2<>"0" or obs2="ON" then obs1="checked" else obs1=""

if request.form<>"" then
	if parentesco="Titular" then
		sqlp="select nome, dtnascimento from qry_funcionarios where chapa='" & rs("chapa") & "' "
		rsc.Open sqlp, ,adOpenStatic, adLockReadOnly
		if rsc.recordcount>0 then nome_bolsista=rsc("nome")	
		if rsc.recordcount>0 then dtnasc=rsc("dtnascimento")
		rsc.close
	end if
end if
%>
<form method="POST" action="bolsa_alteracao.asp" name="form">
<input type="hidden" name="id_bolsa" size="4" value="<%=rs("id_bolsa")%>" >  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Bolsa de Estudo&nbsp;</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Tipo de Bolsa</td>
	<td class=titulo>Funcionário beneficiado</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="tp_bolsa">
<%
sqla="SELECT * from bolsistas_tipo ORDER by descricao "
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rsc("id_tp")=rs("tp_bolsa") then tempt="selected" else tempt=""
%>
		<option value="<%=rsc("id_tp")%>" <%=tempt%>><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select></td>
	<td class=fundo><b><%=rs("chapa")%> - <%=rsnome("nome")%></b></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Parentesco/Tipo</td>
	<td class=titulo>Nome do bolsista</td>
	<td class=titulo>Nascimento</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="parentesco" onChange="javascript:submit()">
          <option value=""></option>
<%
idade=int((now-dtnasc)/365.25)
varpar(0)="Titular"
varpar(1)="Filho"
varpar(2)="Filha"
varpar(3)="Conjuge"
varpar(4)="Companheira/o"

for a=0 to 4
	if parentesco=varpar(a) then tempp="selected" else tempp=""
%>
	<option value="<%=varpar(a)%>" <%=tempp%>><%=varpar(a)%></option>
<%
next
%>
	</select></td>
	<td class=fundo><input type="text" name="nome_bolsista"  size="45" value="<%=nome_bolsista%>"></td>
	<td class=fundo><input type="text" name="dtnasc" size="12" value="<%=dtnasc%>"> (<%=idade%>)</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Situação  </td>
	<td class=titulo>Tipo Curso</td>
	<td class=titulo>Curso     </td>
</tr>
<tr>
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
	<td class=fundo><select size="1" name="tipocurso">
	<option value=""></option>
<%
varcur(0)="Graduação"
varcur(1)="Especialização"
varcur(2)="Mestrado"
varcur(3)="Doutorado"
varcur(4)="Pós-Doutorado"
varcur(5)="Tecnológico"
varcur(6)="Outros"

for a=0 to 6
	if tipocurso=varcur(a) then tempp="selected" else tempp=""
%>
	<option value="<%=varcur(a)%>" <%=tempp%>><%=varcur(a)%></option>
<%
next
%>
	</select></td>

	<td class=fundo><input type="text" name="curso"  size="40" value="<%=curso%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Instituição de ES</td>
	<td class=titulo>Observação </td>
	<td class=titulo>Matrícula  </td>
	<td class=titulo>Comprovante</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="instituicao"  size="15" value="<%=instituicao%>"></td>
	<td class=fundo><input type="text" name="observacao"  size="30" value="<%=observacao%>"></td>
	<td class=fundo><input type="text" name="matricula"  size="10" value="<%=matricula%>"></td>
	<td class=fundo><input type="checkbox" name="comprovante" value="ON" <%=obs1 %>></td>
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
end if
set rs=nothing
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