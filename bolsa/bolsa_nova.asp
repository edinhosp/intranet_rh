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
<title>Inclusão de Bolsa de Estudo</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function chapa1() {	form.chapa.value=form.nome.value;	}
function nome1() {	form.nome.value=form.chapa.value;	}
--></script>
<script language="VBScript">
sub dtnasc_OnChange
	data=document.form.dtnasc.value
	idade=int((now()-cdate(data))/365.25)
	mensagem="Este bolsista excedeu a idade limite para bolsas. Tem " & idade & " anos!"
	if idade>24 then msgbox mensagem,48,"Idade acima do limite"
end sub
</script>
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

if request.form<>"" then
if request.form("bt_salvar")<>"" then
	tudook=1
	'response.write request.form
	if request.form("comprovante")="ON" then compl = -1 else compl = 0
	sql = "INSERT INTO bolsistas (" 
	sql = sql & "tp_bolsa, chapa, parentesco, nome_bolsista, "
	sql = sql & "dtnasc, situacao, tipocurso, curso, instituicao, "
	sql = sql & "observacao, matricula, comprovante "
	sql = sql & ") "
	sql2 = " SELECT "
	sql2=sql2 & " '" & request.form("tp_bolsa") & "', "
	sql2=sql2 & " '" & request.form("chapa") & "', "
	sql2=sql2 & " '" & request.form("parentesco") & "', "
	sql2=sql2 & " '" & request.form("nome_bolsista") & "', "
	if request.form("dtnasc")=""    then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dtnasc")) & "', "
	sql2=sql2 & " '" & request.form("situacao") & "', "
	sql2=sql2 & " '" & request.form("tipocurso") & "', "
	sql2=sql2 & " '" & request.form("curso") & "', "
	sql2=sql2 & " '" & request.form("instituicao") & "', "
	sql2=sql2 & " '" & request.form("observacao") & "', "
	sql2=sql2 & " '" & request.form("matricula") & "', "
	sql2=sql2 & compl & " "

	sql1 = sql & sql2 & ""
	'response.write "<font size='1'>" & sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
end if
else 'request.form=""
end if

'if request.form="" then
if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then

'if request.form("parentesco")<>""    then parentesco=rs("parentesco")       else parentesco=request.form("parentesco")
'if request.form("nome_bolsista")<>"" then nome_bolsista=rs("nome_bolsista") else nome_bolsista=request.form("nome_bolsista")
'if request.form("dtnasc")<>""        then dtnasc=rs("dtnasc")               else dtnasc=request.form("dtnasc")
'if request.form("situacao")<>""      then situacao=rs("situacao")           else situacao=request.form("situacao")
'if request.form("curso")<>""         then curso=rs("curso")                 else curso=request.form("curso")
'if request.form("instituicao")<>""   then instituicao=rs("instituicao")     else instituicao=request.form("instituicao")
'if request.form("tipocurso")<>""     then tipocurso=rs("tipocurso")         else tipocurso=request.form("tipocurso")
'if request.form("observacao")<>""    then observacao=rs("observacao")       else observacao=request.form("observacao")
'if request.form("matricula")<>""     then matricula=rs("matricula")         else matricula=request.form("matricula")
'if request.form("comprovante")<>""   then obs2=rs("comprovante")            else obs2=request.form("comprovante")
tp_bolsa=request.form("tp_bolsa")
parentesco=request.form("parentesco")
nome_bolsista=request.form("nome_bolsista")
dtnasc=request.form("dtnasc")
situacao=request.form("situacao")
curso=request.form("curso")
instituicao=request.form("instituicao")
tipocurso=request.form("tipocurso")
observacao=request.form("observacao")
matricula=request.form("matricula")
obs2=request.form("comprovante")
if obs2="ON" then obs1="checked" else obs1=""

if request.form<>"" then
	if request.form("parentesco")="Titular" then
		sqlp="select nome, dtnascimento from qry_funcionarios where chapa='" & request.form("chapa") & "' "
		rsc.Open sqlp, ,adOpenStatic, adLockReadOnly
		if rsc.recordcount>0 then nome_bolsista=rsc("nome")	
		if rsc.recordcount>0 then dtnasc=rsc("dtnascimento")
		rsc.close
	end if
end if

%>
<form method="POST" action="bolsa_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Inclusão de Bolsa de Estudo</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Tipo de Bolsa</td>
	<td class=titulo>Funcionário beneficiado</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="tp_bolsa">
<%
sqla="SELECT * from bolsistas_tipo ORDER by descricao"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if tp_bolsa=rsc("id_tp") then tempt="selected" else tempt=""
%>
	<option value="<%=rsc("id_tp")%>" <%=tempt%> ><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select>	</td>
	<td class=fundo><select size="1" name="chapa">
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' union select chapa, nome from corporerm.dbo.pfunc where chapa in ('01224','01213') "
if request("chapa")<>"" then sql2=sql2 & "and chapa='" & request("chapa") & "'" else sql2=sql2 & "order by nome"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if request.form("chapa")=rsc("chapa") then tempc="selected" else tempc=""
%>
	<option value="<%=rsc("chapa")%>" <%=tempc%>><%=rsc("nome")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select></td>
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
varpar(0)="Titular":varpar(1)="Filho"
varpar(2)="Filha"  :varpar(3)="Conjuge"
varpar(4)="Companheira/o"
for a=0 to 4
	if parentesco=varpar(a) then tempp="selected" else tempp=""
%>
	<option value="<%=varpar(a)%>" <%=tempp%> ><%=varpar(a)%></option>
<%
next
%>
	</select>
	</td>
	<td class=fundo><input type="text" name="nome_bolsista" size="45" value="<%=nome_bolsista%>" ></td>
	<td class=fundo><input type="text" name="dtnasc" size="12" value="<%=dtnasc%>" ></td>
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
if situacao=rsc("id_sit") then tempt="selected" else tempt=""
%>
	<option value="<%=rsc("id_sit")%>" <%=tempt%> ><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
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
	if request.form("tipocurso")=varcur(a) then tempp="selected" else tempp=""
%>
	<option value="<%=varcur(a)%>" <%=tempp%>><%=varcur(a)%></option>
<%
next
%>
	</select></td>
	<td class=fundo><input type="text" name="curso" size="40" value="<%=request.form("curso")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Instituição de ES</td>
	<td class=titulo>Observação</td>
	<td class=titulo>Matrícula</td>
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
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
else
'rs.close
set rs=nothing
set rsc=nothing
end if
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